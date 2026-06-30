"""Tests for L1 document lifecycle tools."""
import os
import sys
import tempfile

import pytest


DOC_LIFECYCLE_AVAILABLE = sys.platform == "win32"


DOC_LIFECYCLE_AVAILABLE = sys.platform == "win32"


@pytest.mark.skipif(not DOC_LIFECYCLE_AVAILABLE, reason="COM tests only on Windows")
class TestDocumentLifecycle:
    """End-to-end tests for document lifecycle tools."""

    @pytest.fixture(autouse=True)
    def setup(self):
        """Ensure a Word instance exists and a temp file path is available."""
        import asyncio
        from word_mcp.com_manager import get_word_app
        self.app = get_word_app(visible=False)
        self.tmpdir = tempfile.mkdtemp()
        self.loop = asyncio.new_event_loop()
        yield
        self.loop.close()
        # Cleanup: close any open documents
        try:
            while self.app.Documents.Count > 0:
                self.app.Documents(1).Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        """Helper: run an async tool function synchronously."""
        return self.loop.run_until_complete(coro)

    def test_create_and_save_document(self):
        """word_create + word_save: create a doc and save it."""
        from word_mcp.tools.lifecycle import word_create, word_save

        result1 = self._run(word_create())
        assert "已创建" in result1

        doc_path = os.path.join(self.tmpdir, "test_create.docx")
        result2 = self._run(word_save(file_path=doc_path))
        assert "已保存" in result2
        assert os.path.exists(doc_path)

    def test_open_existing_document(self):
        """word_open: open a .docx file."""
        from word_mcp.tools.lifecycle import word_open

        doc_path = os.path.join(self.tmpdir, "test_open.docx")
        doc = self.app.Documents.Add()
        doc.SaveAs(doc_path)
        doc.Close()

        result = self._run(word_open(file_path=doc_path))
        assert "已打开" in result

    def test_export_pdf(self):
        """word_save_as_pdf: export document to PDF."""
        from word_mcp.tools.lifecycle import word_save_as_pdf

        doc = self.app.Documents.Add()
        doc.Content.Text = "Hello PDF"

        pdf_path = os.path.join(self.tmpdir, "test_export.pdf")
        result = self._run(word_save_as_pdf(output_path=pdf_path))
        assert "已导出" in result
        assert os.path.exists(pdf_path)

    def test_close_document(self):
        """word_close: close active document."""
        from word_mcp.tools.lifecycle import word_close

        self.app.Documents.Add()
        assert self.app.Documents.Count == 1

        result = self._run(word_close(save_before_close=False))
        assert "已关闭" in result
        assert self.app.Documents.Count == 0

    def test_get_active_document_info(self):
        """word_get_active_document: return doc info."""
        from word_mcp.tools.lifecycle import word_get_active_document

        doc = self.app.Documents.Add()
        doc.Content.Text = "Test content " * 100

        result = self._run(word_get_active_document())
        assert "当前文档" in result
        assert "页数" in result


@pytest.mark.skipif(not DOC_LIFECYCLE_AVAILABLE, reason="COM tests only on Windows")
class TestEndToEnd:
    """End-to-end scenario: create a complete document."""

    @pytest.fixture(autouse=True)
    def setup(self):
        import asyncio
        from word_mcp.com_manager import get_word_app

        self.app = get_word_app(visible=False)
        self.loop = asyncio.new_event_loop()
        self.tmpdir = tempfile.mkdtemp()
        yield
        self.loop.close()
        try:
            while self.app.Documents.Count > 0:
                self.app.Documents(1).Close(SaveChanges=False)
        except Exception:
            pass

    def _run(self, coro):
        return self.loop.run_until_complete(coro)

    def test_create_resume_document(self):
        """Simulate: agent creates a resume with table + formatted text."""
        from word_mcp.tools.lifecycle import (
            word_create, word_save, word_get_active_document,
        )
        from word_mcp.tools.operations import (
            word_insert_text, word_format_text, word_insert_table,
            word_format_table, word_insert_page_break,
        )

        # 1. Create new document
        self._run(word_create())
        assert self.app.Documents.Count == 1

        # 2. Insert title
        self._run(word_insert_text(text="张三的简历\n", position="start"))
        self._run(word_format_text(range="all", bold=True, font_size=22))

        # 3. Insert personal info table
        data = [
            ["姓名", "张三"],
            ["电话", "138-0000-0000"],
            ["邮箱", "zhangsan@example.com"],
        ]
        self._run(word_insert_table(rows=3, cols=2, data=data))
        self._run(word_format_table(table_index=1, header_row=True))

        # 4. Insert work experience section
        self._run(word_insert_text(text="\n工作经历\n"))

        # 5. Insert page break
        self._run(word_insert_page_break())

        # 6. Save
        doc_path = os.path.join(self.tmpdir, "简历.docx")
        result = self._run(word_save(file_path=doc_path))
        assert "已保存" in result
        assert os.path.exists(doc_path)

        # 7. Verify document has content
        info = self._run(word_get_active_document())
        assert "页数" in info
