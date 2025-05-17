from src.constants import NR_PROJECT, CLIENT_NAME, LOCAL, MARKUP, CONTRACTED_METTER_VALUE, HEADER_START, PROJECT_CODE


def get_excel_metadata(sheet):
    """Extracts project details and header information."""
    metadata = {
        "nr_project": sheet.range(NR_PROJECT).value,
        "client_name": sheet.range(CLIENT_NAME).value,
        "project_code": sheet.range(PROJECT_CODE).value,
        "local": sheet.range(LOCAL).value,
        "markup_factor": sheet.range(MARKUP).value,
        "contracted_m3_value": sheet.range(CONTRACTED_METTER_VALUE).value
    }

    # Get headers from the principal table
    header_start_range = sheet.range(f"{HEADER_START}")
    metadata["headers"] = header_start_range.expand("right").value

    return metadata

