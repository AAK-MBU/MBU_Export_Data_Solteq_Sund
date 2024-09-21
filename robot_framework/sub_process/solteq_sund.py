"""Module"""
import os
import time
import datetime
import pyodbc
import uiautomation as auto


class SolteqSundApp:
    """
    A class to automate interactions with the Solteq Sund application.
    """

    def __init__(self, app_path, username, password, ssn, db_connection_string):
        """
        Initializes the SolteqSundApp object.

        Args:
            app_path (str): Path to the application.
            username (str): Username for login.
            password (str): Password for login.
            ssn (str): SSN for lookup.
        """
        self.app_path = app_path
        self.username = username
        self.password = password
        self.ssn = ssn
        self.connection_string = db_connection_string
        self.app_window = None

    def _is_journal_created(self) -> bool:
        """
        Checks to see if the journal has been stored on the case by SQL lookup.
        """
        current_date = datetime.datetime.now().strftime("%d-%m-%Y")
        filename = f"Udskrift af journal {current_date}.pdf"

        conn = pyodbc.connect(f'{self.connection_string}', autocommit=True)
        crsr = conn.cursor()
        crsr.execute(
            f"""
            SELECT	ds.DocumentId
                    ,ds.entityId
                    ,ds.OriginalFilename
                    ,ds.UniqueFilename
                    ,dss.Document_HistoryId
                    ,dss.Documented
                    ,dss.Decided
                    ,dss.DocumentStoreStatusId
            FROM DocumentStore ds
            JOIN Child c ON c.childId = ds.entityId
            JOIN (
                SELECT DocumentId, MAX(Document_HistoryId) AS MaxDocument_HistoryId
                FROM DocumentStoreStatus
                GROUP BY DocumentId
            ) maxHistory ON ds.DocumentId = maxHistory.DocumentId
            JOIN DocumentStoreStatus dss ON ds.DocumentId = dss.DocumentId AND dss.Document_HistoryId = maxHistory.MaxDocument_HistoryId
            WHERE	c.cpr = '{self.ssn.replace("-", "")}'
                    AND ds.OriginalFilename = '{filename}'
                    AND dss.DocumentStoreStatusId = 1
            ORDER BY Decided DESC
            """
        )
        rows = crsr.fetchall()
        result = []
        columns = [column[0] for column in crsr.description]

        for row in rows:
            result.append(dict(zip(columns, row)))

        if not result:
            bool_result = False
        else:
            bool_result = True

        return bool_result

    def wait_for_control(self, control_type, search_params, search_depth=1, timeout=30):
        """
        Waits for a given control type to become available with the specified search parameters.

        Args:
            control_type: The type of control, e.g., auto.WindowControl, auto.ButtonControl, etc.
            search_params (dict): Search parameters used to identify the control.
                                  The keys must match the properties used in the control type, e.g., 'AutomationId', 'Name'.
            search_depth (int): How deep to search in the user interface.
            timeout (int): How long to wait, in seconds.

        Returns:
            Control: The control object if found, otherwise None.
        """
        end_time = time.time() + timeout
        while time.time() < end_time:
            control = control_type(searchDepth=search_depth, **search_params)
            if control.Exists(0, 0):
                return control
            time.sleep(0.5)
        raise TimeoutError(f"Control with parameters {search_params} was not found within the timeout period.")

    def wait_for_control_to_disappear(self, control_type, search_params, search_depth=1, timeout=30):
        """
        Waits for a given control type to disappear with the specified search parameters.

        Args:
            control_type: The type of control, e.g., auto.WindowControl, auto.ButtonControl, etc.
            search_params (dict): Search parameters used to identify the control.
                                The keys must match the properties used in the control type, e.g., 'AutomationId', 'Name'.
            search_depth (int): How deep to search in the user interface.
            timeout (int): How long to wait, in seconds.

        Returns:
            bool: True if the control disappeared within the timeout period, otherwise False.
        """
        end_time = time.time() + timeout
        while time.time() < end_time:
            control = control_type(searchDepth=search_depth, **search_params)
            if not control.Exists(0, 0):
                return True
            time.sleep(0.5)
        raise TimeoutError(f"Control with parameters {search_params} did not disappear within the timeout period.")

    def start_application(self):
        """
        Starts the application using the specified path.
        """
        os.startfile(self.app_path)

    def login(self):
        """
        Logs into the application by entering the username and password.
        Checks if the login window is open and ready.
        Checks if the main window is opened and ready.
        """
        self.app_window = self.wait_for_control(
            auto.WindowControl,
            {'AutomationId': 'frmLogin'},
            search_depth=2,
        )
        self.app_window.SetFocus()

        username_box = self.app_window.EditControl(AutomationId="textBoxLogin")
        username_box.SendKeys(text=self.username, waitTime=0)

        password_box = self.app_window.EditControl(AutomationId="textBoxPassword")
        password_box.SendKeys(text=self.password, waitTime=0)
        password_box.SendKeys('{ENTER}', waitTime=0)

        self.app_window = self.wait_for_control(
            auto.WindowControl,
            {'AutomationId': 'frmClient'},
            search_depth=2,
            timeout=60
        )

    def open_patient(self):
        """
        Opens the search tab in Solteq Sund and searches for the given patient using SSN.
        """
        self.app_window.SendKeys("{Ctrl}o", waitTime=0)
        self.wait_for_control(
            auto.EditControl,
            {"AutomationId": "TextBoxChildCPR"},
            search_depth=8
        )
        ssn_textbox = self.app_window.EditControl(AutomationId="TextBoxChildCPR")
        ssn_textbox.SetFocus()
        ssn_textbox.SendKeys(self.ssn)
        ssn_textbox.SendKeys("{ENTER}")

        list_item = self.wait_for_control(
            auto.ListItemControl,
            {"Name": f"{self.ssn}"},
            search_depth=12
        )
        list_item.DoubleClick(simulateMove=False, waitTime=0)

        self.wait_for_control(
            auto.TabItemControl,
            {"Name": f"{self.ssn}".replace("-", "")},
            search_depth=4
        )

    def create_journal(self):
        """
        Creates the journal and store it in the documentfolder.
        """
        self.app_window.SendKeys("{Ctrl}{Shift}p", waitTime=0)
        print_journal_window = self.wait_for_control(
            auto.WindowControl,
            {"AutomationId": "frmViewBase"},
            search_depth=2
        )
        print_journal_button = print_journal_window.PaneControl(AutomationId="buttonPrintToDocumentStore")
        print_journal_button.Click(simulateMove=False, waitTime=0)

        self.wait_for_control_to_disappear(
            auto.WindowControl,
            {"AutomationId": "frmViewBase"},
            search_depth=2,
            timeout=60
        )

        if not self._is_journal_created():
            raise ValueError("The document was not found in the SQL database.")
