from word_mcp_codemode_live.core.word_com import undo_record


class _UndoRecord:
    IsRecordingCustomRecord = False

    def __init__(self) -> None:
        self.started: list[str] = []
        self.ended = 0

    def StartCustomRecord(self, name: str) -> None:
        self.started.append(name)
        self.IsRecordingCustomRecord = True

    def EndCustomRecord(self) -> None:
        self.ended += 1
        self.IsRecordingCustomRecord = False


class _App:
    def __init__(self) -> None:
        self.UndoRecord = _UndoRecord()


def test_nested_undo_records_join_the_outer_transaction() -> None:
    app = _App()

    with undo_record(app, "outer"):
        with undo_record(app, "inner one"):
            pass
        with undo_record(app, "inner two"):
            pass

    assert app.UndoRecord.started == ["outer"]
    assert app.UndoRecord.ended == 1
