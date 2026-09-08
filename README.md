# sofficeTest

Simple testbed and python wrappers for some of the implementations python-uno gives to connect and manipulate a Calc workbook.

## Documentation

See the [API documentation](docs/api.md).
Documentation was generated with `pydoc-markdown .pydoc-markdown.yml`

## Testing

macOS note: Recent LibreOffice releases have changed the security and code-signing requirements for the Python runtime bundled inside the LibreOffice application. Which makes it tricky to load uno inside pythonscripts.

There is a workaround here though: https://bugs.documentfoundation.org/show_bug.cgi?id=167752#c17

Use this exported python to run scripts that import uno.

For example:
`/path-to-libreoffice-bundled-python-with-new-signatures/Resoures/python -m unittest`
