# Contributing to ADA Audit Tool

Thank you for considering contributing to the ADA Audit Tool! This document provides guidelines and instructions for contributing to the project.

## Code of Conduct

### Our Standards

- **Be respectful**: Treat everyone with respect and consideration
- **Be inclusive**: Welcome newcomers and help them learn
- **Be collaborative**: Work together to achieve common goals
- **Be professional**: Focus on constructive feedback and solutions

## How to Contribute

### Reporting Bugs

Before creating a bug report:
1. Check the [Issues](https://github.com/yourusername/ada-audit-tool/issues) to see if it's already reported
2. Collect information about the bug:
   - Python version (`python --version`)
   - Operating system and version
   - Steps to reproduce the issue
   - Expected vs. actual behavior
   - Screenshots if applicable

Create a detailed bug report including:
```markdown
## Description
[Clear description of the bug]

## Steps to Reproduce
1. [First step]
2. [Second step]
3. [...]

## Expected Behavior
[What should happen]

## Actual Behavior
[What actually happens]

## Environment
- OS: [e.g., Windows 11, macOS 13.0, Ubuntu 22.04]
- Python Version: [e.g., 3.10.5]
- Package Versions: [paste output of `pip list`]

## Additional Context
[Any other relevant information]
```

### Suggesting Features

Feature requests are welcome! Please provide:
- **Use case**: Describe the problem you're trying to solve
- **Proposed solution**: How should the feature work?
- **Alternatives**: What alternatives have you considered?
- **Impact**: Who would benefit from this feature?

### Pull Requests

#### Development Process

1. **Fork the repository**
   ```bash
   git clone https://github.com/yourusername/ada-audit-tool.git
   cd ada-audit-tool
   ```

2. **Create a virtual environment**
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-bug-fix
   ```

5. **Make your changes**
   - Write clear, documented code
   - Follow the code style guidelines (see below)
   - Add tests if applicable
   - Update documentation as needed

6. **Test your changes**
   - Run the application and verify functionality
   - Test accessibility features (keyboard navigation, screen reader compatibility)
   - Test on multiple operating systems if possible

7. **Commit your changes**
   ```bash
   git add .
   git commit -m "feat: add new feature description"
   # or
   git commit -m "fix: resolve specific bug"
   ```

   Use conventional commit messages:
   - `feat:` New feature
   - `fix:` Bug fix
   - `docs:` Documentation changes
   - `style:` Code style changes (formatting, etc.)
   - `refactor:` Code refactoring
   - `test:` Adding or updating tests
   - `chore:` Maintenance tasks

8. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```

9. **Create a Pull Request**
   - Go to the original repository on GitHub
   - Click "New Pull Request"
   - Select your fork and branch
   - Provide a clear title and description
   - Reference any related issues

#### Pull Request Guidelines

- **One feature per PR**: Keep pull requests focused on a single feature or fix
- **Keep it small**: Smaller PRs are easier to review and merge
- **Update documentation**: Include relevant documentation updates
- **Test thoroughly**: Ensure your changes don't break existing functionality
- **Maintain accessibility**: All UI changes must maintain WCAG 2.1 AA compliance

## Code Style Guidelines

### Python Style

Follow [PEP 8](https://pep8.org/) style guidelines:

```python
# Good
def calculate_attendance_total(student_data, program_code):
    """
    Calculate total attendance for a specific program.
    
    Args:
        student_data (pd.DataFrame): Student attendance data
        program_code (str): Program identifier (e.g., 'Prog_C')
    
    Returns:
        float: Total attendance value
    """
    # Implementation here
    pass

# Bad
def calc(d,p):
    # No docstring, unclear variable names
    pass
```

### Naming Conventions

- **Functions**: Use `snake_case` (e.g., `extract_student_data`)
- **Classes**: Use `PascalCase` (e.g., `ADAAuditGUI`)
- **Constants**: Use `UPPER_SNAKE_CASE` (e.g., `MAX_PROGRAMS`)
- **Variables**: Use descriptive `snake_case` names

### Documentation

All functions must have docstrings:

```python
def find_program_boundaries(data, program_name):
    """
    Find the start and end rows for a program in the dataset.
    
    This function searches through the dataset to locate where a specific
    program's data begins and ends, which is essential for accurate data
    extraction.
    
    Args:
        data (pd.DataFrame): The attendance data to search
        program_name (str): The full program name to search for
    
    Returns:
        tuple: (start_row, end_row) as integers, or (None, None) if not found
    
    Example:
        >>> start, end = find_program_boundaries(df, "Program C Charter Resident")
        >>> print(f"Program data spans rows {start} to {end}")
    """
    # Implementation
    pass
```

### Accessibility Requirements

All UI changes must maintain accessibility:

- **Color contrast**: Text must have 4.5:1 contrast ratio minimum
- **Keyboard navigation**: All functionality accessible via keyboard
- **Focus indicators**: Clear visual indication of focused elements
- **Screen reader support**: Proper labels and announcements
- **Font size**: Minimum 11pt for body text, 12pt for buttons

### Testing Guidelines

While formal tests are being developed, please manually test:

1. **Functionality**: Does the feature work as intended?
2. **Edge cases**: What happens with invalid input?
3. **Accessibility**: Can you use it with keyboard only?
4. **Cross-platform**: Does it work on Windows, macOS, Linux?
5. **Performance**: Does it handle large datasets efficiently?

## Project Structure

```
ada-audit-tool/
├── ADA_Audit_GUI.py              # Main GUI application
├── ADA_Audit_25_26_IMPROVED.py   # Core processing logic
├── ADA_Dashboard_Module.py        # Dashboard generation
├── boundary_settings/             # Configuration files
├── requirements.txt               # Python dependencies
├── README.md                      # Project documentation
├── CONTRIBUTING.md                # This file
├── LICENSE                        # MIT License
└── .gitignore                     # Git ignore rules
```

## Development Setup

### Recommended Tools

- **IDE**: VS Code, PyCharm, or similar with Python support
- **Linter**: flake8 or pylint
- **Formatter**: black (optional but recommended)
- **Git client**: Command line or GUI tool of your choice

### Virtual Environment

Always use a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows
pip install -r requirements.txt
```

## Getting Help

- **Questions**: Open a [Discussion](https://github.com/yourusername/ada-audit-tool/discussions)
- **Bugs**: Create an [Issue](https://github.com/yourusername/ada-audit-tool/issues)
- **Chat**: [Link to Discord/Slack if available]

## Recognition

Contributors will be recognized in:
- Project README.md
- Release notes
- Contributors list

Thank you for helping make the ADA Audit Tool better! 🎉

