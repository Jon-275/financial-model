"""Identify supported state tax rules and their published data vintages."""

from dataclasses import dataclass


WORKBOOK_STATE_POLICY_SCHEMA = 'workbook-state-policy-v1'
WORKBOOK_STATE_POLICY_SCHEMA_CELL = 'B266'
WORKBOOK_CA_POLICY_VERSION_CELL = 'B267'
WORKBOOK_JURISDICTION_ROW = 136
WORKBOOK_JURISDICTION_LABEL = 'State tax jurisdiction'
WORKBOOK_JURISDICTION_ERROR = 'Check state tax jurisdiction'


@dataclass(frozen=True, slots=True)
class StatePolicy:
    """Versioned rules for a jurisdiction, with separate income and payroll sources."""

    jurisdiction: str
    version: str
    published_income_year: int
    published_sdi_year: int
    income_source: str
    sdi_source: str


CALIFORNIA_POLICY = StatePolicy(
    jurisdiction='CA',
    version='ca-individual-v1',
    published_income_year=2026,
    published_sdi_year=2027,
    income_source='https://www.ftb.ca.gov/about-ftb/newsroom/tax-news/2026/10.html',
    sdi_source='https://edd.ca.gov/en/Payroll_Taxes/Rates_and_Withholding/',
)


def resolve_state_policy(jurisdiction: object, version: object) -> StatePolicy:
    """Resolve explicit canonical values without a fallback for unknown input."""
    if type(jurisdiction) is not str or jurisdiction != CALIFORNIA_POLICY.jurisdiction:
        raise ValueError('Unsupported state tax jurisdiction.')
    if type(version) is not str or version != CALIFORNIA_POLICY.version:
        raise ValueError('Unsupported state tax policy version.')
    return CALIFORNIA_POLICY
