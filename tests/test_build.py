"""Tests for build pipeline improvements."""
import json
import os
import pytest


class TestMetadataValidation:
    def test_missing_styles_raises(self, tmp_dir):
        """build() should raise clear error when metadata.json lacks 'styles'."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="styles"):
            _validate_metadata({"id": "test", "sections": []})

    def test_missing_multiple_fields(self, tmp_dir):
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="styles.*sections|sections.*styles"):
            _validate_metadata({"id": "test"})

    def test_valid_metadata_passes(self, tmp_dir):
        from hwpx_engine.build import _validate_metadata
        # Should not raise
        _validate_metadata({"id": "test", "styles": {}, "sections": []})

    def test_missing_id_raises(self, tmp_dir):
        """build() should raise clear error when metadata.json lacks 'id'."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="id"):
            _validate_metadata({"styles": {}, "sections": []})

    def test_empty_metadata_raises(self, tmp_dir):
        """Empty metadata dict should report all missing fields."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="id"):
            _validate_metadata({})

    def test_error_message_lists_all_missing(self, tmp_dir):
        """Error message should name all missing fields, not just the first."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError) as exc_info:
            _validate_metadata({"id": "test"})
        msg = str(exc_info.value)
        assert "styles" in msg
        assert "sections" in msg

    def test_context_appears_in_error_message(self, tmp_dir):
        """Error message should include the context string passed in."""
        from hwpx_engine.build import _validate_metadata
        with pytest.raises(ValueError, match="register"):
            _validate_metadata({"id": "test", "sections": []}, context="register")


class TestRegisterTemplateHardening:
    def _make_src(self, tmp_dir, name="foo", with_summary=True):
        import json, os
        src = os.path.join(tmp_dir, "src_" + name)
        os.makedirs(src, exist_ok=True)
        meta = {"id": name, "styles": {}, "sections": []}
        if with_summary:
            meta["display_name"] = name.title()
            meta["summary"] = "one-line purpose"
        with open(os.path.join(src, "metadata.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f)
        with open(os.path.join(src, "template.hwpx"), "wb") as f:
            f.write(b"PK\x03\x04stub")
        return src

    def test_rejects_invalid_id(self, tmp_dir, monkeypatch):
        import sys
        from pathlib import Path
        # Get the build module from sys.modules (not the build function)
        from hwpx_engine import registry
        build_module = sys.modules['hwpx_engine.build']
        reg_path = Path(tmp_dir) / "registered"
        monkeypatch.setattr(build_module, "GLOBAL_REGISTERED", reg_path)
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg_path)
        src = self._make_src(tmp_dir, "foo")
        with pytest.raises(registry.InvalidTemplateIdError):
            build_module.register_template("BAD NAME", src)

    def test_refuses_overwrite_without_force(self, tmp_dir, monkeypatch):
        import sys
        from pathlib import Path
        from hwpx_engine import registry
        build_module = sys.modules['hwpx_engine.build']
        reg = Path(tmp_dir) / "registered"
        monkeypatch.setattr(build_module, "GLOBAL_REGISTERED", reg)
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)
        src = self._make_src(tmp_dir, "foo")
        build_module.register_template("foo", src)
        with pytest.raises(registry.TemplateAlreadyExistsError):
            build_module.register_template("foo", src)

    def test_force_allows_overwrite(self, tmp_dir, monkeypatch):
        """force=True must actually replace the existing content, not just no-op."""
        import sys
        from pathlib import Path
        from hwpx_engine import registry
        build_module = sys.modules['hwpx_engine.build']
        reg = Path(tmp_dir) / "registered"
        monkeypatch.setattr(build_module, "GLOBAL_REGISTERED", reg)
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)

        # First registration with marker "v1"
        src_v1 = self._make_src(tmp_dir, "foo_v1")
        meta_v1 = json.loads(Path(src_v1, "metadata.json").read_text(encoding="utf-8"))
        meta_v1["version_marker"] = "v1"
        Path(src_v1, "metadata.json").write_text(json.dumps(meta_v1), encoding="utf-8")
        build_module.register_template("foo", src_v1)
        assert json.loads((reg / "foo" / "metadata.json").read_text(encoding="utf-8"))["version_marker"] == "v1"

        # Second registration with marker "v2" using force=True
        src_v2 = self._make_src(tmp_dir, "foo_v2")
        meta_v2 = json.loads(Path(src_v2, "metadata.json").read_text(encoding="utf-8"))
        meta_v2["version_marker"] = "v2"
        Path(src_v2, "metadata.json").write_text(json.dumps(meta_v2), encoding="utf-8")
        build_module.register_template("foo", src_v2, force=True)

        # Verify the overwritten content is v2, not v1
        final = json.loads((reg / "foo" / "metadata.json").read_text(encoding="utf-8"))
        assert final["version_marker"] == "v2", f"force did not actually overwrite: {final}"

    def test_missing_summary_still_allowed(self, tmp_dir, monkeypatch):
        """summary is optional — registration should succeed without it.
        list_templates() will mark it incomplete separately."""
        import sys
        from pathlib import Path
        from hwpx_engine import registry
        build_module = sys.modules['hwpx_engine.build']
        reg = Path(tmp_dir) / "registered"
        monkeypatch.setattr(build_module, "GLOBAL_REGISTERED", reg)
        monkeypatch.setattr(registry, "GLOBAL_REGISTERED", reg)
        src = self._make_src(tmp_dir, "foo", with_summary=False)
        build_module.register_template("foo", src)  # should not raise
        assert (reg / "foo" / "metadata.json").exists()
