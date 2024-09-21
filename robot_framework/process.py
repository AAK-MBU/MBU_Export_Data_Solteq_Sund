"""This module contains the main process of the robot."""
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from robot_framework.sub_process.solteq_sund import SolteqSundApp
from robot_framework import config


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    app_obj = SolteqSundApp(
        app_path=rf"{config.APP_PATH}",
        username=f"{orchestrator_connection.get_credential("solteq_sund").username}",
        password=f"{orchestrator_connection.get_credential("solteq_sund").password}",
        ssn="",
        db_connection_string=orchestrator_connection.get_constant("solteq_sund_db_connstr").value
    )

    app_obj.start_application()
    app_obj.login()
    app_obj.open_patient()
    app_obj.create_journal()


if __name__ == "__main__":
    oc = OrchestratorConnection.create_connection_from_args()
    process(oc)
