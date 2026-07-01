import json
import pytest
import sys

@pytest.mark.skipif(sys.platform != "win32", reason="COM tests only on Windows")
class TestStructureAndLocator:
    @pytest.fixture(autouse=True)
    def setup(self):
        from word_mcp.com_manager import get_word_app
        from word_mcp.tools.lifecycle import word_create
        self.app = get_word_app(visible=False)
        word_create()
        self.doc = self.app.ActiveDocument
        yield
        try:
            self.doc.Close(SaveChanges=False)
        except Exception:
            pass

    def test_document_structure(self):
        from word_mcp.tools.text import word_insert_text
        from word_mcp.tools.structure import word_get_document_structure
        
        word_insert_text("Title\n", position="end")
        word_insert_text("Some text\n", position="end")
        
        # apply style to title
        from word_mcp.tools.text import word_apply_style
        word_apply_style(range_spec={"type": "paragraph", "index": 1}, style_name="Heading 1")
        
        res = word_get_document_structure(include_paragraphs=True)
        assert "Title" in res
        data = json.loads(res)
        assert len(data) >= 2
        assert data[0]["type"] == "heading"
        
    def test_delete_and_replace_target(self):
        from word_mcp.tools.text import word_insert_text, word_delete_target, word_replace_target
        word_insert_text("Para 1\n", position="end")
        word_insert_text("Para 2\n", position="end")
        
        word_replace_target(target={"type": "paragraph", "index": 1}, new_text="Replaced 1\n")
        assert "Replaced 1" in self.doc.Paragraphs(1).Range.Text
        
        word_delete_target(target={"type": "paragraph", "index": 2})
        assert "Para 2" not in self.doc.Content.Text
