from virmetrix.application.session import ExamMode, ExamSession, SessionStatus


def test_practice_session_no_timer():
    s = ExamSession(mode=ExamMode.PRACTICE, exercise_ids=["a", "b"])
    s.start()
    assert s.status == SessionStatus.RUNNING
    assert s.remaining_seconds is None
    s.tick()
    assert s.status == SessionStatus.RUNNING
    s.finish()
    assert s.status == SessionStatus.FINISHED
    # idempotent
    s.finish()
    assert s.status == SessionStatus.FINISHED


def test_exam_session_countdown_and_finish():
    s = ExamSession(mode=ExamMode.EXAM, exercise_ids=["a"], duration_seconds=2)
    s.start()
    assert s.remaining_seconds == 2
    s.tick()
    assert s.remaining_seconds == 1
    s.tick()
    assert s.remaining_seconds == 0
    assert s.status == SessionStatus.FINISHED


def test_pause_resume():
    s = ExamSession(mode=ExamMode.EXAM, exercise_ids=["a"], duration_seconds=5)
    s.start()
    s.pause()
    assert s.status == SessionStatus.PAUSED
    s.tick()
    assert s.remaining_seconds == 5
    s.resume()
    assert s.status == SessionStatus.RUNNING


def test_navigation():
    s = ExamSession(mode=ExamMode.PRACTICE, exercise_ids=["a", "b", "c"])
    assert s.current_exercise_id == "a"
    assert s.next_exercise() is True
    assert s.current_exercise_id == "b"
    assert s.prev_exercise() is True
    assert s.current_exercise_id == "a"
    assert s.prev_exercise() is False