from word_mcp_codemode_live.word.equations import to_unicode_math


def test_unicode_math_prefers_longest_command() -> None:
    assert to_unicode_math(r"\iint_A + \int_B") == "∬_A + ∫_B"


def test_unicode_math_does_not_replace_command_prefixes() -> None:
    assert to_unicode_math(r"\infinite + \infty") == r"\infinite + ∞"
