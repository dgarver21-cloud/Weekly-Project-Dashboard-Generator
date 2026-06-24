from __future__ import annotations

import json
import hashlib
from datetime import date, datetime
from pathlib import Path

import streamlit as st

from dashboard_defaults import current_report_date
from generate_dashboard import build_dashboard_bytes, dashboard_data_from_dict
from parse_adi_export import AdiExportError, build_dashboard_data, read_csv_text
from parse_clarizen_plan import ClarizenPlanError, parse_clarizen_workbook
from parse_meeting_transcript import parse_meeting_transcripts


APP_DIR = Path(__file__).resolve().parent
BASE_DATA_PATH = APP_DIR / "sample_inputs" / "dashboard_sample.json"
TEMPLATE_PATH = APP_DIR / "templates" / "weekly-status-template.pptx"


@st.cache_data
def load_base_dashboard() -> dict:
    return json.loads(BASE_DATA_PATH.read_text(encoding="utf-8"))


def parse_display_date(value: str, fallback: date) -> date:
    for date_format in ("%B %d, %Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(value, date_format).date()
        except (TypeError, ValueError):
            continue
    return fallback


def build_project_data(
    base_data: dict,
    project_name: str,
    project_manager: str,
    project_sponsor: str,
    report_date: date,
    go_live_date: date | None,
    project_status: str,
) -> dict:
    data = dict(base_data)
    data["project_summary"] = {
        "project_name": project_name.strip(),
        "project_manager": project_manager.strip(),
        "project_sponsor": project_sponsor.strip(),
        "report_date": report_date.strftime("%B %d, %Y"),
        "go_live_date": (
            go_live_date.strftime("%B %d, %Y") if go_live_date else ""
        ),
        "project_status": project_status.strip(),
    }
    return data


def build_editor_rows(all_items: list[dict], selected_items: list[dict]) -> list[dict]:
    return [
        {
            "Include on Dashboard": index < len(selected_items),
            "RAID Type": item.get("type", ""),
            "Description": item.get("description", ""),
            "Priority": item.get("priority", ""),
            "Status": item.get("status", ""),
            "Assigned": item.get("assigned", ""),
            "Due Date": item.get("due_date", ""),
        }
        for index, item in enumerate(all_items)
    ]


def apply_edited_raid_rows(
    original_items: list[dict], edited_rows: list[dict]
) -> list[dict]:
    included_items = []
    for index, row in enumerate(edited_rows):
        if not bool(row.get("Include on Dashboard")):
            continue
        original = original_items[index] if index < len(original_items) else {}
        updated = dict(original)
        updated.update(
            {
                "type": str(row.get("RAID Type", "")).strip(),
                "description": str(row.get("Description", "")).strip(),
                "priority": str(row.get("Priority", "")).strip(),
                "status": str(row.get("Status", "")).strip(),
                "assigned": str(row.get("Assigned", "")).strip(),
                "due_date": str(row.get("Due Date", "")).strip(),
            }
        )
        description_lines = updated["description"].splitlines()
        if description_lines and description_lines[0].strip():
            updated["title"] = description_lines[0].strip()
        included_items.append(updated)
    return included_items


def editor_records(edited_table) -> list[dict]:
    if hasattr(edited_table, "to_dict"):
        return edited_table.to_dict("records")
    return list(edited_table)


def build_deliverable_editor_rows(deliverables: list[dict]) -> list[dict]:
    return [
        {
            "Include on Dashboard": bool(
                item.get("recommended", index < 5)
            ),
            "Name": item.get("name", ""),
            "Owner": item.get("owner", ""),
            "Start Date": item.get("start_date", item.get("start", "")),
            "End Date": item.get("end_date", item.get("end", "")),
            "State": item.get("state", item.get("status", "Active")),
            "Ranking Reason": item.get("ranking_reason", ""),
        }
        for index, item in enumerate(deliverables)
    ]


def apply_edited_deliverable_rows(
    original_items: list[dict], edited_rows: list[dict]
) -> list[dict]:
    included_items = []
    for index, row in enumerate(edited_rows):
        if not bool(row.get("Include on Dashboard")):
            continue
        original = original_items[index] if index < len(original_items) else {}
        updated = dict(original)
        updated.update(
            {
                "name": str(row.get("Name", "")).strip(),
                "owner": str(row.get("Owner", "")).strip(),
                "start_date": str(row.get("Start Date", "")).strip(),
                "end_date": str(row.get("End Date", "")).strip(),
                "state": str(row.get("State", "")).strip() or "Active",
                "ranking_reason": str(row.get("Ranking Reason", "")).strip(),
            }
        )
        included_items.append(updated)
    return included_items


def edited_deliverable_warnings(deliverables: list[dict]) -> list[str]:
    warnings = []
    for item in deliverables:
        name = item.get("name", "") or "Unnamed deliverable"
        if not item.get("start_date"):
            warnings.append(f"{name}: missing Start Date; no Gantt bar will be drawn.")
        if not item.get("end_date"):
            warnings.append(f"{name}: missing End Date; no Gantt bar will be drawn.")
    return warnings


def apply_transcript_narrative(suggestion_key: str) -> None:
    st.session_state["project_status_narrative"] = st.session_state.get(
        suggestion_key, ""
    ).strip()


def show_transcript_items(title: str, items: list) -> None:
    st.markdown(f"**{title}**")
    if items:
        for item in items:
            if isinstance(item, dict):
                st.write(f"- {item.get('text', '')}")
            else:
                st.write(f"- {item}")
    else:
        if title == "Blockers or concerns":
            st.caption(
                "No clear project blockers were identified from the uploaded "
                "transcripts."
            )
        else:
            st.caption("No items identified.")


def show_transcript_details(title: str, items: list[dict]) -> None:
    if not items:
        return
    st.markdown(f"**{title}**")
    for item in items:
        st.write(item.get("text", ""))
        st.caption(
            f"Source: {item.get('source_file', 'Unknown')} | "
            f"Confidence: {item.get('confidence_score', 0)} | "
            f"{item.get('reason', '')}"
        )
        raw_text = item.get("raw_text", "")
        if raw_text and raw_text != item.get("text", ""):
            st.caption(f"Raw extraction: {raw_text}")


def main() -> None:
    st.set_page_config(
        page_title="Weekly Project Dashboard Generator",
        page_icon=":bar_chart:",
        layout="wide",
    )
    st.title("Weekly Project Dashboard Generator")
    st.caption(
        "Upload ADI and Clarizen exports, optionally review a transcript-based "
        "status draft, and download the completed PowerPoint."
    )

    try:
        base_data = load_base_dashboard()
    except (OSError, json.JSONDecodeError) as error:
        st.error(f"Could not load the sample dashboard settings: {error}")
        st.stop()

    default_report_date = current_report_date()

    st.subheader("Project Information")
    left, right = st.columns(2)
    with left:
        project_name = st.text_input("Project Name", value="")
        project_manager = st.text_input("Project Manager", value="")
        project_sponsor = st.text_input("Project Sponsor", value="")
        report_date = st.date_input("Report Date", value=default_report_date)
    with right:
        go_live_date = st.date_input("Go-Live Date", value=None)
        if "project_status_narrative" not in st.session_state:
            st.session_state["project_status_narrative"] = ""
        project_status = st.text_area(
            "Project Status narrative",
            height=150,
            key="project_status_narrative",
        )

    blank_project_fields = [
        label
        for label, value in (
            ("Project Name", project_name),
            ("Project Manager", project_manager),
            ("Project Sponsor", project_sponsor),
            ("Go-Live Date", go_live_date),
            ("Project Status narrative", project_status),
        )
        if value is None or (isinstance(value, str) and not value.strip())
    ]
    if blank_project_fields:
        st.warning(
            "Recommended Project Information field(s) are blank and will appear "
            "blank in the dashboard: " + ", ".join(blank_project_fields) + "."
        )

    project_data = build_project_data(
        base_data,
        project_name,
        project_manager,
        project_sponsor,
        report_date,
        go_live_date,
        project_status,
    )

    st.subheader("Client Logo")
    client_logo_file = st.file_uploader(
        "Upload Client Logo",
        type=["png", "jpg", "jpeg"],
        help="Optional. The logo appears in the bottom-left of the PowerPoint.",
    )
    st.caption(
        "PNG with a transparent background is recommended. If no logo is "
        "uploaded, the bottom-left logo area remains blank."
    )
    if client_logo_file is not None:
        st.info(f"Uploaded client logo: {client_logo_file.name}")

    st.subheader("ADI List")
    uploaded_file = st.file_uploader(
        "Upload the exported ADI List CSV",
        type=["csv"],
        help="Use the CSV export from the ADI List. No live SharePoint connection is used.",
    )

    st.subheader("Clarizen Project Plan")
    clarizen_file = st.file_uploader(
        "Upload Clarizen Project Plan Excel Export",
        type=["xlsx", "xls"],
        help=(
            "Optional. The Clarizen Excel export populates the Current & Upcoming "
            "Deliverables Gantt chart, not the RAID table."
        ),
    )

    st.subheader("Meeting Transcript")
    transcript_files = st.file_uploader(
        "Upload Meeting Transcript(s)",
        type=["txt", "vtt", "docx"],
        accept_multiple_files=True,
        help=(
            "Optional. One or more transcripts can be combined into a Project "
            "Status draft and review suggestions. They do not update RAID or "
            "deliverables."
        ),
    )

    if transcript_files:
        adi_context_titles = []
        if uploaded_file is not None:
            try:
                context_csv = uploaded_file.getvalue().decode("utf-8-sig")
                context_rows = read_csv_text(context_csv)
                adi_context_titles = [
                    row.get("Action Items", "").strip()
                    for row in context_rows
                    if row.get("Action Items", "").strip()
                ]
            except (UnicodeDecodeError, AdiExportError):
                pass

        clarizen_context_names = []
        if clarizen_file is not None:
            try:
                context_deliverables, _ = parse_clarizen_workbook(
                    clarizen_file.getvalue(),
                    today=date.today(),
                    reference_date=report_date,
                    file_name=clarizen_file.name,
                )
                clarizen_context_names = [
                    item.get("name", "").strip()
                    for item in context_deliverables
                    if item.get("name", "").strip()
                ]
            except ClarizenPlanError:
                pass

        project_context = {
            "project_name": project_name.strip(),
            "project_status_narrative": project_status.strip(),
            "adi_titles": adi_context_titles,
            "clarizen_deliverable_names": clarizen_context_names,
        }
        transcript_data = [
            (transcript_file.getvalue(), transcript_file.name)
            for transcript_file in transcript_files
        ]
        transcript_result = parse_meeting_transcripts(
            transcript_data,
            project_context=project_context,
        )

        if transcript_result is not None:
            context_bytes = json.dumps(
                project_context, sort_keys=True
            ).encode("utf-8")
            fingerprint_data = bytearray(context_bytes)
            for transcript_bytes, transcript_name in transcript_data:
                fingerprint_data.extend(transcript_name.encode("utf-8"))
                fingerprint_data.extend(transcript_bytes)
            fingerprint = hashlib.sha256(fingerprint_data).hexdigest()[:12]
            suggestion_key = f"transcript_suggestion_{fingerprint}"
            suggested_narrative = transcript_result[
                "suggested_project_status_narrative"
            ]
            for warning in transcript_result["transcript_warnings"]:
                st.warning(warning)

            if suggested_narrative:
                if suggestion_key not in st.session_state:
                    st.session_state[suggestion_key] = suggested_narrative
                st.caption(
                    "Review and edit this draft before applying it. The Project "
                    "Status field remains manually editable after it is applied."
                )
                st.text_area(
                    "Suggested Project Status narrative",
                    height=130,
                    key=suggestion_key,
                )
                st.button(
                    "Use suggested transcript narrative",
                    on_click=apply_transcript_narrative,
                    args=(suggestion_key,),
                    disabled=not bool(st.session_state[suggestion_key].strip()),
                )

            st.info(
                "Transcript suggestions are for PM review only and do not update "
                "the official ADI/RAID or Clarizen sections."
            )

            with st.expander("Transcript review", expanded=True):
                show_transcript_items(
                    "Accomplishments", transcript_result["accomplishments"]
                )
                show_transcript_items(
                    "Upcoming focus areas",
                    transcript_result["upcoming_focus_areas"],
                )
                show_transcript_items(
                    "Blockers or concerns",
                    transcript_result["blockers_or_concerns"],
                )
                show_transcript_items(
                    "Decisions needed", transcript_result["decisions_needed"]
                )
                show_transcript_items(
                    "Possible RAID/ADI items for review",
                    transcript_result["possible_raid_items_for_review"],
                )
                st.caption(
                    "Possible RAID/ADI items are for PM awareness only and are not "
                    "added to the official RAID table."
                )
                show_transcript_items(
                    "Transcript warnings", transcript_result["transcript_warnings"]
                )
            detail_categories = (
                ("Accomplishments", "accomplishments"),
                ("Upcoming focus areas", "upcoming_focus_areas"),
                ("Blockers or concerns", "blockers_or_concerns"),
                ("Decisions needed", "decisions_needed"),
                (
                    "Possible RAID/ADI items for review",
                    "possible_raid_items_for_review",
                ),
                ("Low-confidence items ignored", "low_confidence_items"),
            )
            with st.expander("Show transcript extraction details", expanded=False):
                for title, category in detail_categories:
                    show_transcript_details(title, transcript_result[category])

    clarizen_warnings = []
    if clarizen_file is not None:
        try:
            parsed_deliverables, clarizen_warnings = parse_clarizen_workbook(
                clarizen_file.getvalue(),
                today=date.today(),
                reference_date=report_date,
                file_name=clarizen_file.name,
            )
        except ClarizenPlanError as error:
            st.error(str(error))
            return

        st.caption(
            "Only Active tasks with Level 2 or higher are shown. The five most "
            "relevant current/upcoming deliverables are included by default."
        )
        deliverable_rows = build_deliverable_editor_rows(parsed_deliverables)
        edited_deliverable_table = st.data_editor(
            deliverable_rows,
            use_container_width=True,
            hide_index=True,
            disabled=["Ranking Reason"],
            column_config={
                "Include on Dashboard": st.column_config.CheckboxColumn(
                    "Include on Dashboard"
                ),
                "Name": st.column_config.TextColumn("Name", width="large"),
                "Owner": st.column_config.TextColumn("Owner"),
                "Start Date": st.column_config.TextColumn(
                    "Start Date", help="Use YYYY-MM-DD."
                ),
                "End Date": st.column_config.TextColumn(
                    "End Date", help="Use YYYY-MM-DD."
                ),
                "State": st.column_config.TextColumn("State"),
                "Ranking Reason": st.column_config.TextColumn(
                    "Ranking Reason",
                    help="Why this deliverable was recommended and ranked.",
                ),
            },
            key=f"deliverable_editor_{clarizen_file.name}_{clarizen_file.size}",
        )
        edited_deliverables = apply_edited_deliverable_rows(
            parsed_deliverables, editor_records(edited_deliverable_table)
        )
    else:
        edited_deliverables = list(project_data.get("deliverables", []))
        st.info(
            "No Clarizen file uploaded. The dashboard will use the existing "
            "sample/default deliverables."
        )

    deliverable_count = len(edited_deliverables)
    edit_date_warnings = edited_deliverable_warnings(edited_deliverables)
    if deliverable_count > 5:
        st.warning(
            "More than 5 deliverables are selected. The Gantt section is designed "
            "for 5 rows and may become crowded. Reduce the selection to 5."
        )
    elif deliverable_count == 0:
        st.warning("Select at least one deliverable before generating the PowerPoint.")
    else:
        st.info(
            f"{deliverable_count} deliverable(s) will be included in the Gantt chart."
        )

    if clarizen_file is not None:
        with st.expander(
            f"Clarizen PM review warnings ({len(clarizen_warnings)})",
            expanded=bool(clarizen_warnings),
        ):
            if clarizen_warnings:
                for warning in clarizen_warnings:
                    st.warning(warning)
            else:
                st.success("No Clarizen PM review warnings.")
    if edit_date_warnings:
        with st.expander(
            f"Selected deliverable date warnings ({len(edit_date_warnings)})",
            expanded=True,
        ):
            for warning in edit_date_warnings:
                st.warning(warning)

    if uploaded_file is None:
        st.info("Upload an ADI List CSV to preview RAID items and generate the dashboard.")
        return

    try:
        csv_text = uploaded_file.getvalue().decode("utf-8-sig")
        rows = read_csv_text(csv_text)
        selected_dashboard, warnings, active_count = build_dashboard_data(
            rows, project_data, limit=5, today=date.today()
        )
        all_active_dashboard, _, _ = build_dashboard_data(
            rows, project_data, limit=max(1, active_count), today=date.today()
        )
    except UnicodeDecodeError:
        st.error("The CSV could not be read. Re-save it as CSV UTF-8 in Excel.")
        return
    except AdiExportError as error:
        st.error(str(error))
        return

    raid_items = selected_dashboard["raid_items"]
    all_active_items = all_active_dashboard["raid_items"]
    st.success(
        f"Selected {len(raid_items)} of {active_count} active ADI items for the dashboard."
    )

    st.subheader("Review and Edit RAID Items")
    st.caption(
        "The ranked top five items are included by default. Edit values directly "
        "or change which items are included."
    )
    editor_rows = build_editor_rows(all_active_items, raid_items)
    edited_table = st.data_editor(
        editor_rows,
        use_container_width=True,
        hide_index=True,
        disabled=["RAID Type"],
        column_config={
            "Include on Dashboard": st.column_config.CheckboxColumn(
                "Include on Dashboard",
                help="Uncheck an item to leave it out of the PowerPoint.",
            ),
            "RAID Type": st.column_config.TextColumn("RAID Type"),
            "Description": st.column_config.TextColumn(
                "Description",
                width="large",
                help="The first line is the title. Remaining lines are regular text.",
            ),
            "Priority": st.column_config.SelectboxColumn(
                "Priority",
                options=["Critical", "High", "Normal", "Low", "Unassigned"],
            ),
            "Status": st.column_config.SelectboxColumn(
                "Status",
                options=["In Progress", "On Hold", "Ready for UAT", "Client Review"],
            ),
            "Assigned": st.column_config.TextColumn("Assigned"),
            "Due Date": st.column_config.TextColumn(
                "Due Date", help="Use YYYY-MM-DD when possible."
            ),
        },
        key=f"raid_editor_{uploaded_file.name}_{uploaded_file.size}",
    )
    edited_rows = editor_records(edited_table)
    edited_raid_items = apply_edited_raid_rows(all_active_items, edited_rows)
    included_count = len(edited_raid_items)

    if included_count > 5:
        st.warning(
            "More than 5 RAID items are selected. The slide is designed for 5 "
            "items and may become crowded. Reduce the selection to 5 before generating."
        )
    elif included_count == 0:
        st.warning("Select at least one RAID item before generating the PowerPoint.")
    else:
        st.info(f"{included_count} RAID item(s) will be included in the PowerPoint.")

    with st.expander(
        f"PM review warnings ({len(warnings)})", expanded=bool(warnings)
    ):
        if warnings:
            for warning in warnings:
                st.warning(warning)
        else:
            st.success("No PM review warnings.")

    if st.button(
        "Generate PowerPoint",
        type="primary",
        disabled=(
            included_count == 0
            or included_count > 5
            or deliverable_count == 0
            or deliverable_count > 5
        ),
    ):
        try:
            dashboard_json = dict(selected_dashboard)
            dashboard_json["project_summary"] = project_data["project_summary"]
            dashboard_json["raid_items"] = edited_raid_items
            dashboard_json["deliverables"] = edited_deliverables
            dashboard_data, parsed_report_date = dashboard_data_from_dict(
                dashboard_json
            )
            pptx_bytes = build_dashboard_bytes(
                TEMPLATE_PATH,
                dashboard_data,
                parsed_report_date,
                client_logo_bytes=(
                    client_logo_file.getvalue()
                    if client_logo_file is not None
                    else None
                ),
            )
        except (OSError, ValueError) as error:
            st.error(f"Could not generate the PowerPoint: {error}")
            return

        st.session_state["generated_pptx"] = pptx_bytes
        st.session_state["generated_filename"] = (
            f"{project_name.strip() or 'project'}-weekly-status-report.pptx"
        )
        st.success("PowerPoint generated successfully.")

    if "generated_pptx" in st.session_state:
        st.download_button(
            "Download PowerPoint",
            data=st.session_state["generated_pptx"],
            file_name=st.session_state["generated_filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    main()
