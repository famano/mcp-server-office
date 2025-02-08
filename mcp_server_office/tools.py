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
        "Make text replacements in specified paragraphs of a docx file. "
        "Accepts a list of edits with paragraph index and search/replace pairs. "
        "Each edit operates on a single paragraph and preserves the formatting of the first run. "
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
                "description": "Sequence of edits to apply to specific paragraphs.",
                "items": {
                    "type": "object",
                    "properties": {
                        "paragraph_index": {
                            "type": "integer",
                            "description": "0-based index of the paragraph to edit. tips: whole table is count as one paragraph."
                        },
                        "search": {
                            "type": "string",
                            "description": (
                                "Text to find within the specified paragraph. "
                                "The search is performed only within the target paragraph. "
                                "Escape line break when you input multiple lines."
                            )
                        },
                        "replace": {
                            "type": "string",
                            "description": (
                                "Text to replace the search string with. "
                                "The formatting of the first run in the paragraph will be applied to the entire replacement text. "
                                "Empty string represents deletion. "
                                "Escape line break when you input multiple lines."
                            )
                        }
                    },
                    "required": ["paragraph_index", "search", "replace"]
                }
            }
        },
        "required": ["path", "edits"]
    }
)
