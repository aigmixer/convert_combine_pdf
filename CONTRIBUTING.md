# Contributing to Convert & Merge to PDF

Thank you for considering contributing to this project!

## How to Contribute

### Reporting Issues

- Check existing issues before creating a new one
- Include Windows version, PowerShell version, and tool versions
- Provide error messages and steps to reproduce

### Suggesting Features

- Open an issue with the "enhancement" label
- Describe the use case and expected behavior
- Consider implementation complexity

### Code Contributions

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Make your changes
4. Test thoroughly on Windows 10/11
5. Commit with clear messages
6. Push and create a Pull Request

## Development Guidelines

### Code Style

- Use 4-space indentation
- Add comments for complex logic
- Follow existing PowerShell conventions
- Keep functions focused and small

### Testing

Test with:
- Multiple file types simultaneously
- Large files (>10MB)
- Files with special characters in names
- Files from different directories
- Edge cases (single file, 50+ files)

### Error Handling

- Provide clear error messages
- Don't fail silently
- Clean up temp files even on error
- Log useful debugging information

## Areas for Improvement

- Additional file format support (HTML, RTF, etc.)
- GUI wrapper with progress indicators
- Configurable output location/naming
- PDF compression/optimization options
- Multi-language support
- Drag-and-drop interface

## License

By contributing, you agree that your contributions will be licensed under GPL v3.
