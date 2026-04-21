from __future__ import annotations  # allows using list[str]

from dataclasses import dataclass   
from typing import Optional          # for None fields
import pandas as pd                  


@dataclass
class Email:
    email_number: int        # which email in the thread this is
    subject: str             # email subject line
    body: str                # the full email body text
    sender: str              # who sent the email
    recipients: list[str]    # list of recipients


@dataclass
class Scenario:
    scenario_id: str                  # unique ID for this scenario
    scenario_type: str                # type code from the sheet
    emails: list[Email]               # all emails in this scenario which is now sorted by email number
    success_criteria: Optional[str]   # what correct response looks like
    puzzle_summary: Optional[str]     # basic description of the scenario


def _clean_str(value) -> Optional[str]:
    """Return stripped string or None for NaN/blank values."""
    if pd.isna(value):       # pd.isna catches NaN, None, and float('nan') from pandas
        return None
    s = str(value).strip()   
    return s if s else None  # return None if the string was all whitespace


def load_scenarios(path: str) -> list[Scenario]:
    df = pd.read_excel(path, dtype=str)  # read every cell as a string

    df.columns = [col.replace("\u200b", "").strip() for col in df.columns] # Strip zero-width spaces 

    df = df[df["Scenario Type"].notna() | df["Sender"].notna()].copy() # Drop blank spacer rows
    df = df.reset_index(drop=True)  # renumber rows after dropping rows

    
    df = df[df["Scenario ID"].apply(lambda v: _clean_str(v) not in ("Example:", "xxx"))] # This one removed the "Example:" header row and any placeholder "xxx" rows but may not be nessiary later so
    df = df.reset_index(drop=True)  

    df["Scenario Type"] = df["Scenario Type"].ffill() # ffill() copies the last seen value downward to fill those blanks being if Multi-email scenarios leave Scenario Type blank on rows after the first one

    df["_group"] = (df["Scenario Type"] != df["Scenario Type"].shift()).cumsum() # Create a group number that increments every time Scenario Type changes so for example two dif T01 get different group IDs

    scenarios: list[Scenario] = []  # holds the final list of Scenario objects

    for _, group in df.groupby("_group", sort=True):  # iterate over each scenario's rows together
        group = group.copy()  # copy so we can safely add columns without affecting the original dataframe but unsure if nedeed

        scenario_type = _clean_str(group.iloc[0]["Scenario Type"]) or ""  # grab type from the first row

        scenario_id = _clean_str(group.iloc[0].get("Scenario ID")) or scenario_type # Use the Scenario ID column if it has a value or if not just use the type code as the ID

        
        group["_email_num"] = pd.to_numeric(group["Email #"], errors="coerce")  # # Convert Email # to a number so we can sort correctly cuz Excel reads it as a string, this one turns non-numeric to NaN
        group = group.sort_values("_email_num", na_position="last")  # ranks them from numbered emails first to unnumbered last but unsure if thats how it should play out

        emails: list[Email] = []           # collects Email objects for this scenario
        success_criteria: Optional[str] = None  
        puzzle_summary: Optional[str] = None   

        for _, row in group.iterrows():  # iterate over each email row in this scenario

            if not _clean_str(row.get("Sender")) and not _clean_str(row.get("Subject")): # Skip rows with no sender and no subject
                continue

            recipients_raw = _clean_str(row.get("Recipient(s)"))  # raw string like "CEO, V"
            recipients = (
                [r.strip() for r in recipients_raw.split(",") if r.strip()]  # split on comma, strip spaces
                if recipients_raw   # only split if there's actually a value
                else []             # otherwise use an empty list
            )

            email_num_raw = row.get("Email #")  # raw value from the cell
            try:
                email_num = int(float(email_num_raw)) if pd.notna(email_num_raw) else 0 # removes decimal because execl be giving floats like 1.0 insted of 1
            except (ValueError, TypeError):
                email_num = 0  # if conversion fails for any reason, default to 0

            emails.append(
                Email(
                    email_number=email_num,
                    subject=_clean_str(row.get("Subject")) or "",  # empty string if blank
                    body=_clean_str(row.get("Body")) or "",        # placeholders like {date} kept raw(as was)
                    sender=_clean_str(row.get("Sender")) or "",
                    recipients=recipients,
                )
            )

            # Success Criteria and Puzzle Summary only appear on one row per scenario but im unsure if thats how we want it
            # It grabs the success criteria from whichever row has it and stores it on the whole scenario insted of each individual email.

            sc = _clean_str(row.get("Success Criteria"))
            if sc:                      # only overwrite if this row actually has a value
                success_criteria = sc
            ps = _clean_str(row.get("Puzzle Summary"))
            if ps:
                puzzle_summary = ps

        if not emails:   # skip over spacers so not empty scienarios
            continue

        scenarios.append(
            Scenario(
                scenario_id=scenario_id,
                scenario_type=scenario_type,
                emails=emails,
                success_criteria=success_criteria,
                puzzle_summary=puzzle_summary,
            )
        )

    return scenarios


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else "Emails.xlsx"  # accept path as CLI arg or default
    scenarios = load_scenarios(path)
    print(f"Loaded {len(scenarios)} scenarios\n")
    for s in scenarios:
        print(f"[{s.scenario_type}] id={s.scenario_id!r}  criteria={s.success_criteria!r}")
        for e in s.emails:
            print(f"  Email #{e.email_number}: {e.subject!r}")
            print(f"    From: {e.sender!r}  To: {e.recipients}")
            #print(f"    Body: {e.body[:80]!r}{'...' if len(e.body) > 80 else ''}")  # this basiccaly cuts it off at a limit ending it in ... if you want to use it
            print(f"    Body: {e.body}")
        print()
