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


class TestListTemplates:
    def test_empty_when_no_dir(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", tmp_path / "registered")
        result = registry.list_templates()
        assert result == []
        assert (tmp_path / "registered").is_dir()

    def test_lists_ok_template(self, tmp_path, monkeypatch):
        import json
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        t = reg / "foo"
        t.mkdir(parents=True)
        (t / "metadata.json").write_text(json.dumps({
            "id": "foo",
            "display_name": "Foo Template",
            "summary": "Short purpose line.",
            "styles": {}, "sections": [],
        }), encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert len(result) == 1
        entry = result[0]
        assert entry["id"] == "foo"
        assert entry["display_name"] == "Foo Template"
        assert entry["summary"] == "Short purpose line."
        assert entry["status"] == "ok"
        assert "path" in entry

    def test_incomplete_when_summary_missing(self, tmp_path, monkeypatch):
        import json
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        t = reg / "foo"
        t.mkdir(parents=True)
        (t / "metadata.json").write_text(json.dumps({
            "id": "foo",
            "display_name": "Foo",
            "styles": {}, "sections": [],
        }), encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert result[0]["status"] == "incomplete"
        assert "summary" in result[0]["missing_fields"]

    def test_invalid_when_json_broken(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        t = reg / "broken"
        t.mkdir(parents=True)
        (t / "metadata.json").write_text("{not json", encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert result[0]["status"] == "invalid"
        assert "error" in result[0]

    def test_skips_dirs_without_metadata(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        (reg / "no_meta").mkdir(parents=True)
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert result == []

    def test_skips_hidden_dirs(self, tmp_path, monkeypatch):
        import json
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        hidden = reg / ".trash"
        hidden.mkdir(parents=True)
        (hidden / "metadata.json").write_text(json.dumps({"id": ".trash"}), encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert result == []

    def test_sorted_by_id(self, tmp_path, monkeypatch):
        import json
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        for name in ["zebra", "alpha", "middle"]:
            t = reg / name
            t.mkdir(parents=True)
            (t / "metadata.json").write_text(json.dumps({
                "id": name, "display_name": name, "summary": "",
                "styles": {}, "sections": [],
            }), encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        ids = [e["id"] for e in result]
        assert ids == ["alpha", "middle", "zebra"]

    def test_empty_string_display_name_treated_as_missing(self, tmp_path, monkeypatch):
        """display_name=\"\" should be treated as missing even though the key exists."""
        import json
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        t = reg / "foo"
        t.mkdir(parents=True)
        (t / "metadata.json").write_text(json.dumps({
            "id": "foo",
            "display_name": "",
            "summary": "has summary",
            "styles": {}, "sections": [],
        }), encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert result[0]["status"] == "incomplete"
        assert "display_name" in result[0]["missing_fields"]

    def test_non_dict_metadata_is_invalid(self, tmp_path, monkeypatch):
        """metadata.json that is valid JSON but not a dict → invalid."""
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        t = reg / "foo"
        t.mkdir(parents=True)
        (t / "metadata.json").write_text('["not", "a", "dict"]', encoding="utf-8")
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        result = registry.list_templates()
        assert result[0]["status"] == "invalid"
        assert "not dict" in result[0]["error"]


class TestUnregisterTemplate:
    def _make_template(self, reg, name="foo"):
        import json
        t = reg / name
        t.mkdir(parents=True)
        (t / "metadata.json").write_text(json.dumps({
            "id": name, "styles": {}, "sections": [],
        }), encoding="utf-8")
        (t / "template.hwpx").write_bytes(b"PK\x03\x04stub")
        return t

    def test_nonexistent_raises(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", tmp_path / "registered")
        monkeypatch.setattr(registry, "TRASH_DIR", tmp_path / ".trash")
        with pytest.raises(registry.TemplateNotFoundError):
            registry.unregister_template("missing")

    def test_invalid_id_raises(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", tmp_path / "registered")
        with pytest.raises(registry.InvalidTemplateIdError):
            registry.unregister_template("BAD ID")

    def test_default_backup_moves_to_trash(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        trash = tmp_path / ".trash"
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)
        monkeypatch.setattr(registry, "TRASH_DIR", trash)
        self._make_template(reg, "foo")

        registry.unregister_template("foo")

        assert not (reg / "foo").exists()
        # Exactly one backup dir in trash, whose name starts with "foo_"
        entries = list(trash.iterdir())
        assert len(entries) == 1
        assert entries[0].name.startswith("foo_")

    def test_force_removes_permanently(self, tmp_path, monkeypatch):
        from hwpx_engine import registry
        reg = tmp_path / "registered"
        trash = tmp_path / ".trash"
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)
        monkeypatch.setattr(registry, "TRASH_DIR", trash)
        self._make_template(reg, "foo")

        registry.unregister_template("foo", backup=False)

        assert not (reg / "foo").exists()
        assert not trash.exists() or not list(trash.iterdir())
