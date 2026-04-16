"""Tests for template registry API."""
import pytest


class TestIdValidation:
    def test_valid_id_passes(self):
        from hwpx_engine.registry import validate_template_id
        # Should not raise
        validate_template_id("gri_policy_research_a4")

    def test_uppercase_rejected(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError, match="lowercase"):
            validate_template_id("GRI_A4")

    def test_spaces_rejected(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError, match="lowercase|characters"):
            validate_template_id("my template")

    def test_special_chars_rejected(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError):
            validate_template_id("my-template!")

    def test_reserved_name_rejected(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError, match="reserved"):
            validate_template_id("registry")

    def test_trash_reserved(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError, match="reserved"):
            validate_template_id(".trash")

    def test_empty_rejected(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError):
            validate_template_id("")

    def test_too_long_rejected(self):
        from hwpx_engine.registry import validate_template_id, InvalidTemplateIdError
        with pytest.raises(InvalidTemplateIdError, match=r"64|length"):
            validate_template_id("a" * 65)
