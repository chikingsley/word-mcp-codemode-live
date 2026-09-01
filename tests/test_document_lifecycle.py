import json
from pathlib import Path

import pytest

from word_mcp_codemode_live.tools.files import (
    copy_document,
    create_document,
    list_available_documents,
)


@pytest.mark.asyncio
async def test_create_and_copy_require_explicit_overwrite(tmp_path: Path) -> None:
    source = tmp_path / "source.DOCX"
    created = json.loads(await create_document(str(source)))
    assert created["success"] is True
    assert source.exists()

    refused_create = json.loads(await create_document(str(source)))
    assert "already exists" in refused_create["error"]

    destination = tmp_path / "destination.docx"
    copied = json.loads(await copy_document(str(source), str(destination)))
    assert copied["success"] is True
    refused_copy = json.loads(await copy_document(str(source), str(destination)))
    assert "already exists" in refused_copy["error"]


@pytest.mark.asyncio
async def test_list_documents_is_sorted_and_case_insensitive(tmp_path: Path) -> None:
    await create_document(str(tmp_path / "zeta.docx"))
    await create_document(str(tmp_path / "Alpha.DOCX"))

    result = json.loads(await list_available_documents(str(tmp_path)))

    assert result["count"] == 2
    assert [Path(item["path"]).name for item in result["documents"]] == [
        "Alpha.DOCX",
        "zeta.docx",
    ]
