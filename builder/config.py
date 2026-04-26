"""
Audit configuration dataclass.

Defines the inputs that drive workbook generation. Used by the
orchestrator (builder.workbook) and passed to every tab builder
that needs user-configurable values.

Validation lives here so that invalid configurations are rejected
before any workbook generation begins. The validator is called
explicitly via AuditConfig.validate(); construction does not
auto-validate (this preserves the option to construct partial
configurations during testing).

The defaults defined here match the v4 reference workbook so that
tests and production code can use the same dataclass while
preserving golden-file parity for the default-configured case.
"""

from dataclasses import dataclass, field
from datetime import date


# Reference defaults matching the v4 golden file.
# These values produce a workbook that matches tests/golden_files/v4_reference.xlsx
# when no user inputs override them.
DEFAULT_KICKOFF_DATE = date(2026, 4, 1)
DEFAULT_PLANNING_WEEKS = 4
DEFAULT_FIELDWORK_WEEKS = 16
DEFAULT_REPORTING_WEEKS = 4
DEFAULT_ON_TARGET_BUFFER = 1
DEFAULT_HOURS_PER_HOLIDAY = 8

# Hard limits per ADR-010
MIN_PHASE_WEEKS = 1
MAX_TOTAL_WEEKS = 52
MIN_ON_TARGET_BUFFER = 0
MAX_HOURS_PER_HOLIDAY = 24


@dataclass
class AuditConfig:
    """
    Configuration for a single audit workbook generation.

    Attributes:
        kickoff_date: The audit's kickoff / project launch date.
            Used as the anchor from which all phase dates are computed.
        planning_weeks: Number of weeks for the Planning phase. Minimum 1.
        fieldwork_weeks: Number of weeks for the Fieldwork phase. Minimum 1.
        reporting_weeks: Number of weeks for the Reporting phase. Minimum 1.
        on_target_buffer: Weeks before end of Fieldwork for the On-Target
            meeting. Default 1, matching v4 reference.
        hours_per_holiday: Hours deducted per closure or skeleton-crew day.
            Default 8.

    Validation rules (enforced by validate()):
        - All phase weeks must be at least MIN_PHASE_WEEKS (1).
        - Sum of phase weeks must not exceed MAX_TOTAL_WEEKS (52).
        - on_target_buffer must be at least MIN_ON_TARGET_BUFFER (0).
        - hours_per_holiday must be between 0 and MAX_HOURS_PER_HOLIDAY (24).

    Validation rules deferred to Sprint 3.5 or later:
        - Kickoff-date constraints per ADR-009 (Friday warning, holiday shift,
          past-date limits). These are UX behaviors better handled by the
          Streamlit UI than the dataclass.
    """

    kickoff_date: date = field(default_factory=lambda: DEFAULT_KICKOFF_DATE)
    planning_weeks: int = DEFAULT_PLANNING_WEEKS
    fieldwork_weeks: int = DEFAULT_FIELDWORK_WEEKS
    reporting_weeks: int = DEFAULT_REPORTING_WEEKS
    on_target_buffer: int = DEFAULT_ON_TARGET_BUFFER
    hours_per_holiday: int = DEFAULT_HOURS_PER_HOLIDAY

    @property
    def total_weeks(self) -> int:
        """Sum of all three phase week counts."""
        return self.planning_weeks + self.fieldwork_weeks + self.reporting_weeks

    def validate(self) -> list[str]:
        """
        Check this configuration against the validation rules.

        Returns:
            List of error messages. Empty list means the config is valid.
        """
        errors = []

        if self.planning_weeks < MIN_PHASE_WEEKS:
            errors.append(
                f"Planning weeks must be at least {MIN_PHASE_WEEKS} "
                f"(got {self.planning_weeks})"
            )
        if self.fieldwork_weeks < MIN_PHASE_WEEKS:
            errors.append(
                f"Fieldwork weeks must be at least {MIN_PHASE_WEEKS} "
                f"(got {self.fieldwork_weeks})"
            )
        if self.reporting_weeks < MIN_PHASE_WEEKS:
            errors.append(
                f"Reporting weeks must be at least {MIN_PHASE_WEEKS} "
                f"(got {self.reporting_weeks})"
            )

        if self.total_weeks > MAX_TOTAL_WEEKS:
            errors.append(
                f"Total audit duration ({self.total_weeks} weeks) exceeds "
                f"maximum {MAX_TOTAL_WEEKS} weeks"
            )

        if self.on_target_buffer < MIN_ON_TARGET_BUFFER:
            errors.append(
                f"On-target buffer must be at least {MIN_ON_TARGET_BUFFER} "
                f"(got {self.on_target_buffer})"
            )

        if self.hours_per_holiday < 0 or self.hours_per_holiday > MAX_HOURS_PER_HOLIDAY:
            errors.append(
                f"Hours per holiday must be between 0 and {MAX_HOURS_PER_HOLIDAY} "
                f"(got {self.hours_per_holiday})"
            )

        return errors

    def is_valid(self) -> bool:
        """Convenience: True if validate() returns no errors."""
        return len(self.validate()) == 0


def default_config() -> AuditConfig:
    """
    Convenience constructor for the v4 reference defaults.

    Used by tests to produce a config matching the golden file, and
    callable from production code if a fallback is ever needed.
    """
    return AuditConfig()
