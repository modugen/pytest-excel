import pytest


@pytest.fixture(scope="session")
def ifc_file(request):
    path = f"{request.param}.ifc"
    return path
