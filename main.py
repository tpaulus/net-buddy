import sys
from copy import copy
from typing import Optional

import npyscreen
import gspread
from datetime import datetime

from gspread.utils import rowcol_to_a1, Dimension

# Google Sheets Setup
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1OopTFBHSx3TWb1ANjdAHr-rpfO-7ImbwpIpwLJ8oLXY/edit?usp=sharing"

gc = gspread.oauth(
    credentials_filename='client_secrets.json',
    authorized_user_filename='.authorized_user.json',
    flow=gspread.auth.local_server_flow,
    scopes=["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/spreadsheets"]
)

now = datetime.now()

sheet = gc.open_by_url(SPREADSHEET_URL).worksheet(str(now.year))
all_rows = sheet.get_all_values()

def get_current_week_col() -> int:
    month_name = now.strftime("%B")
    month_column: int = all_rows[0].index(month_name)

    for week_column in range(month_column, len(all_rows[1])):
        if int(all_rows[1][week_column]) >= now.day:
            return week_column

    raise RuntimeError("Cannot determine current week colum")

current_week_col_0idx = get_current_week_col()
current_week_col_1idx = current_week_col_0idx + 1

def get_member_name(callsign: str) -> Optional[str]:
    for row in all_rows:
        if row[2] == callsign:
            name = row[0]
            position = row[1]

            if position:
                return f"{callsign.upper()} - {name} ({position})"
            return f"{callsign.upper()} - {name}"

    return None


def is_checked_in(callsign: str) -> bool:
    for row in all_rows:
        if row[2] == callsign:
            return row[current_week_col_0idx] == "X"

    return False


def checkin_member(callsign: str):
    if is_checked_in(callsign):
        # Nothing to do, already checked-in
        return

    for i, row in enumerate(all_rows, start=1):
        if row[2] == callsign:
            sheet.cell(i, current_week_col_1idx).value = "X"

def add_row_for_new_operator(name: str, call: str):
    # Get last operator row
    last_member_idx = 0
    for i in range(len(all_rows) -1 , 0, -1):
        if all_rows[i][0] != "" and "Total" not in all_rows[i][0]:
            last_member_idx = i + 1
            break

    # Add a new row
    new_op_idx = last_member_idx + 1
    sheet.insert_row([], index=new_op_idx, inherit_from_before=True)
    sheet.copy_range(f"{last_member_idx}:{last_member_idx}", f"{new_op_idx}:{new_op_idx}")

    # Set Operator Name & Call
    row = [name, "", call]
    row.extend(["" for _ in range(len(row), current_week_col_0idx)]) # A blank value for all the previous weeks of the year
    row.append("X") # And finally an X for this week

    sheet.update([row], f"A{new_op_idx}:{rowcol_to_a1(new_op_idx, len(row)+1)}", True, Dimension.rows)


class MainMenu(npyscreen.FormBaseNew):
    def create(self):
        self.add(npyscreen.TitleFixedText, name="W7AW Net Buddy", value="", editable=False)
        self.add(npyscreen.ButtonPress, name="Check In", when_pressed_function=self.goto_early_checkins)
        self.add(npyscreen.ButtonPress, name="Roll Call", when_pressed_function=self.goto_roll_call)
        self.add(npyscreen.ButtonPress, name="Exit", when_pressed_function=self.exit_application)

    def goto_early_checkins(self):
        self.parentApp.switchForm("CHECKIN")

    def goto_roll_call(self):
        self.parentApp.switchForm("ROLL_CALL")

    @staticmethod
    def exit_application():
        sys.exit(0)


class EarlyCheckins(npyscreen.FormBaseNew):
    def create(self):
        self.call_sign = self.add(npyscreen.TitleText, name="Enter Call Sign:")
        self.add(npyscreen.ButtonPress, name="Check In", when_pressed_function=self.check_in)
        self.add(npyscreen.ButtonPress, name="Back to Main Menu", when_pressed_function=self.back_to_main)

    def check_in(self):
        call_sign = self.call_sign.value.strip().upper()
        if not call_sign:
            npyscreen.notify_confirm("Please enter a call sign.", title="Error")
            return

        try:
            cell = sheet.find(call_sign)
            if cell:
                sheet.update_cell(cell.row, current_week_col_1idx, "X")
                npyscreen.notify_confirm(f"{get_member_name(call_sign)} checked in.", title="Success")
            else:
                add_new_member = npyscreen.notify_yes_no(f"{call_sign} not found. Would you like to add them as a new member?", title='popup')

                if add_new_member:
                    self.parentApp.getForm('NEW_OPERATOR').value = copy(self.call_sign.value).upper()
                    self.parentApp.switchForm('NEW_OPERATOR')

            self.call_sign.value = None
            self.call_sign.focus = True
            self.display()
        except gspread.exceptions.APIError as e:
            npyscreen.notify_confirm(f"Error: {e}", title="API Error")

    def back_to_main(self):
        self.editing = False
        self.parentApp.setNextForm("MAIN")

class NewOperator(npyscreen.ActionForm):
    def create(self):
        self.value = None
        self.call_sign = self.add(npyscreen.TitleText, name = "Call Sign: ")
        self.op_name   = self.add(npyscreen.TitleText, name = "Name:")

    def beforeEditing(self):
        self.name = "New Operator"
        self.call_sign.value = copy(self.value)
        self.op_name.value = ''
        self.op_name.focus = True


    def on_ok(self):
        if not self.call_sign.value:
            npyscreen.notify_confirm("Please enter a call sign.", title="Error")
            return
        if not self.op_name.value:
            npyscreen.notify_confirm("Please enter the operator's name.", title="Error")
            return

        add_row_for_new_operator(self.op_name.value, self.call_sign.value)

        npyscreen.notify_confirm(f"Added {self.op_name.value} ({self.call_sign.value}) to the log", title="Success!")
        self.parentApp.switchFormPrevious()

    def on_cancel(self):
        self.parentApp.switchFormPrevious()



class MemberList(npyscreen.MultiSelect):

    def display_value(self, vl):
        member_name = get_member_name(str(vl))
        if member_name:
            return member_name
        return vl


class RollCall(npyscreen.FormBaseNew):
    def create(self):
        self.add(npyscreen.TitleFixedText, name="Roll Call", value="", editable=False)
        self.members = self.add(MemberList, max_height=(10 if self.lines < 25 else 25))

        self.add(npyscreen.ButtonPress, name="Save & Return to Main Menu", when_pressed_function=self.back_to_main)

    def beforeEditing(self):
        self.members.values = []  # List of options
        self.members.value = []  # Index of selected options

        global all_rows
        all_rows = sheet.get_all_values()

        num_rows = len(all_rows)
        for i, row in enumerate(range(2, num_rows)):
            name = all_rows[row][0]
            if "Do not announce" in name:
                # We have reached the end of the roll-call list
                break

            call_sign = all_rows[row][2]
            if not call_sign:
                continue

            self.members.values.append(call_sign)
            if is_checked_in(call_sign):
                self.members.value.append(len(self.members.values) - 1)

    def back_to_main(self):
        self.editing = False

        # Save checkins
        sheet.update(
            [["X"] if i in self.members.value else [""] for i in range(len(self.members.values))],
f"{rowcol_to_a1(3, current_week_col_1idx)}:{rowcol_to_a1(3 + len(self.members.values), current_week_col_1idx)}")

        # Return
        self.parentApp.setNextForm("MAIN")


class RadioNetApp(npyscreen.NPSAppManaged):
    def onStart(self):
        self.addForm("MAIN", MainMenu)
        self.addForm("CHECKIN", EarlyCheckins)
        self.addForm("NEW_OPERATOR", NewOperator)
        self.addForm("ROLL_CALL", RollCall)


if __name__ == "__main__":
    app = RadioNetApp()
    app.run()
