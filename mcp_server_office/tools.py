from mcp import types

READ_DOCX = types.Tool(
    name="read_docx",
    description=(
        "Read complete contents of a docx file including tables and images."
        "Use this tool when you want to read file endswith '.docx'."
        "Paragraphs are separated with two line breaks."
        "This tool convert images into placeholder [Image]."
        "[delete: xxx] and [insert: xxx] means tracking changes of file."
    ),
    inputSchema={
        "type": "object",
                "properties": {
                "path": {
                        "type": "string",
                        "description": "Absolute path to target file",
                    }
                },
        "required": ["path"]
    }
)

WRITE_DOCX = types.Tool(
    name="write_docx",
    description=(
        "Create a new docx file with given content."
        "Editing exisiting docx file with this tool is not recomended."
    ),
    inputSchema={
        "type": "object",
        "properties": {
            "path": {
                "type": "string",
                "description": "Absolute path to target file. It should be under your current working directory.",
            },
            "content": {
                "type": "string",
                "description": (
                    "Content to write to the file. Two line breaks in content represent new paragraph."
                    "Table should starts with [Table], and separated with '|'."
                    "Escape line break when you input multiple lines."
                ),
            }
        },
        "required": ["path", "content"]
    }
)

EDIT_DOCX = types.Tool(
    name="edit_docx",
    description=(
        "Make multiple text replacements in a docx file. Accepts a list of search/replace pairs "
        "and applies them sequentially. Since this tool is intended to edit a single part of document,"
        "each search should matches exact part of document. Note each search matches only once."
        "Returns a git-style diff showing the changes made. Only works within allowed directories."
    ),
    inputSchema={
        "type": "object",
        "properties": {
            "path": {
                "type": "string",
                "description": "Absolute path to file to edit. It should be under your current working directory."
            },
            "edits": {
                "type": "array",
                "description": "Sequence of edit.",
                "items": {
                    "type": "object",
                    "properties": {
                        "search": {
                            "type": "string",
                            "description": (
                                "search string to find single part of the document."
                                "This should match exact part of document. Search string should unique in document and concise."
                                "Note search string matches only once."
                                "Escape line break when you input multiple lines."
                            )
                        },
                        "replace": {
                            "type": "string",
                            "description": (
                                "replacement of search seach string. Two line breaks in content represent new paragraph."
                                "Table should starts with [Table], and separated with '|'."
                                "Empty string replesents deletion."
                                "Escape line break when you input multiple lines."
                            )
                        }
                    },
                    "required": ["search", "replace"]
                }
            }
        },
        "required": ["path", "edits"]
    }
)