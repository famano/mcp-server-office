import os
import pytest
from mcp_server_office.office import validate_path, read_docx, write_docx, edit_docx, extract_table_text
from docx import Document
from docx.table import Table

@pytest.fixture
def sample_docx():
    """Create a sample docx file for testing."""
    path = "test_sample.docx"
    doc = Document()
    doc.add_paragraph("Hello World")
    
    # Add table
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "A1"
    table.cell(0, 1).text = "B1"
    table.cell(1, 0).text = "A2"
    table.cell(1, 1).text = "B2"
    
    doc.add_paragraph("Goodbye World")
    doc.save(path)
    yield path
    # Cleanup
    if os.path.exists(path):
        os.remove(path)

@pytest.mark.asyncio
async def test_validate_path(sample_docx):
    """Test path validation."""
    assert await validate_path(os.path.abspath(sample_docx)) == True
    with pytest.raises(ValueError):
        await validate_path(sample_docx) 
    with pytest.raises(ValueError):
        await validate_path("nonexistent.docx")

@pytest.mark.asyncio
async def test_read_docx(sample_docx):
    """Test reading docx file."""
    content = await read_docx(os.path.abspath(sample_docx))
    assert "Hello World" in content
    assert "Goodbye World" in content
    assert "[Table]" in content
    assert "A1 | B1" in content
    assert "A2 | B2" in content

    with pytest.raises(ValueError):
        await read_docx("nonexistent.docx")

@pytest.mark.asyncio
async def test_write_docx():
    """Test writing docx file."""
    test_path = "test_write.docx"
    test_content = "Test Paragraph\n\n[Table]\nC1 | D1\nC2 | D2\n\nFinal Paragraph"
    
    try:
        await write_docx(test_path, test_content)
        assert os.path.exists(test_path)
        
        # Verify content
        doc = Document(test_path)
        paragraphs = [p.text for p in doc.paragraphs if p.text]
        assert "Test Paragraph" in paragraphs
        assert "Final Paragraph" in paragraphs
        
        # Verify table
        table = doc.tables[0]
        assert table.cell(0, 0).text == "C1"
        assert table.cell(0, 1).text == "D1"
        assert table.cell(1, 0).text == "C2"
        assert table.cell(1, 1).text == "D2"
    
    finally:
        if os.path.exists(test_path):
            os.remove(test_path)

@pytest.mark.asyncio
async def test_edit_docx(sample_docx):
    abs_sample_docx = os.path.abspath(sample_docx)
    
    """Test editing docx file."""
    # Test single edit
    result = await edit_docx(abs_sample_docx, [{"search": "Hello", "replace": "Hi"}])
    assert "Hello World" in result["original"]
    assert "Hi World" in result["modified"]
    
    # Test multiple edits
    result = await edit_docx(abs_sample_docx, [
        {"search": "Hi", "replace": "Hellow"},
        {"search": "World", "replace": "Everyone"}
    ])
    assert "Hi World" in result["original"]
    assert "Hellow Everyone" in result["modified"]
    
    # Test non-existent text
    with pytest.raises(ValueError) as exc_info:
        await edit_docx(abs_sample_docx, [{"search": "NonexistentText", "replace": "NewText"}])
    assert "Search text not found: NonexistentText" in str(exc_info.value)
    
    # Test multiple edits with one non-existent
    with pytest.raises(ValueError) as exc_info:
        await edit_docx(abs_sample_docx, [
            {"search": "Hello", "replace": "Hi"},
            {"search": "NonexistentText", "replace": "NewText"}
        ])
    assert "Search text not found: NonexistentText" in str(exc_info.value)
    
    # Test invalid file
    with pytest.raises(ValueError):
        await edit_docx("nonexistent.docx", [{"search": "Hello", "replace": "Hi"}])

def test_extract_table_text():
    """Test table text extraction."""
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "X1"
    table.cell(0, 1).text = "Y1"
    table.cell(1, 0).text = "X2"
    table.cell(1, 1).text = "Y2"
    
    result = extract_table_text(table)
    assert "X1 | Y1" in result
    assert "X2 | Y2" in result
