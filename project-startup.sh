#!/bin/bash

# Set the base directory
BASE_DIR="$HOME/Documents/ms_pg_replicator"

# Print header
echo "=========================================="
echo "Setting up MS PG Replicator Project Structure"
echo "=========================================="
echo "Base directory: $BASE_DIR"
echo ""

# Check if base directory exists
if [ ! -d "$BASE_DIR" ]; then
    echo "ERROR: Directory $BASE_DIR does not exist!"
    echo "Please create it first or check the path."
    exit 1
fi

echo "✓ Base directory verified"
echo ""

# Create directory structure
echo "Creating directory structure..."

# Main directories
mkdir -pv "$BASE_DIR/.github/workflows"
mkdir -pv "$BASE_DIR/docs"
mkdir -pv "$BASE_DIR/ms_pg_replicator"
mkdir -pv "$BASE_DIR/tests"
mkdir -pv "$BASE_DIR/scripts"

echo ""
echo "✓ Directory structure created"
echo ""

# Create essential files
echo "Creating essential project files..."

# Create .gitignore (Python-specific)
cat > "$BASE_DIR/.gitignore" << 'EOF'
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
env/
venv/
ENV/
env.bak/
venv.bak/
*.egg-info/
*.egg
.eggs/
dist/
build/
develop-eggs/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
*.log

# Virtual Environment
.venv/
venv/
ENV/

# IDE
.vscode/
.idea/
*.swp
*.swo
*~

# OS
.DS_Store
Thumbs.db
desktop.ini

# Project specific
*.db
*.sqlite
*.accdb
*.mdb
*.ldb
*.laccdb
config.local.py
.env

# Testing
.pytest_cache/
.coverage
htmlcov/
.tox/
.mypy_cache/

# Logs and temporary files
logs/
*.log
temp/
tmp/
EOF

# Create CHANGELOG.md
cat > "$BASE_DIR/CHANGELOG.md" << 'EOF'
# Changelog

All notable changes to the MS PG Replicator project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Initial project structure
- Core replication functionality (planned)
- Access to PostgreSQL migration support (planned)

### Changed
- N/A

### Fixed
- N/A

### Removed
- N/A

## [0.1.0] - 2026-05-23
### Added
- Project initialization
EOF

# Create LICENSE (MIT License - standard for open source)
cat > "$BASE_DIR/LICENSE" << 'EOF'
MIT License

Copyright (c) 2025 MS PG Replicator Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
EOF

# Create pyproject.toml (modern Python project config)
cat > "$BASE_DIR/pyproject.toml" << 'EOF'
[build-system]
requires = ["setuptools>=61.0", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "ms-pg-replicator"
version = "0.1.0"
description = "Microsoft Access to PostgreSQL database replication tool"
authors = [
    {name = "Ron OHara", email = "ronohara@duck.com"},
]
readme = "README.md"
requires-python = ">=3.8"
license = {text = "MIT"}
classifiers = [
    "Development Status :: 3 - Alpha",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.8",
    "Programming Language :: Python :: 3.9",
    "Programming Language :: Python :: 3.10",
    "Programming Language :: Python :: 3.11",
    "Programming Language :: Python :: 3.12",
]

dependencies = [
    "pyodbc>=5.0.0",
    "psycopg2-binary>=2.9.0",
    "pyyaml>=6.0",
    "dotenv>=0.9.9",
    "mmh3>=5.2.1",
    "pip>=26.1.1",
    "psycopg2>=2.9.12",
    "python-dateutil>=2.9.0.post0",
    "python-dotenv>=1.2.2",
    "pywin32>=311",
    "PyYAML>=6.0.3",
    "schedule>=1.2.2",
]

[project.optional-dependencies]
dev = [
    "pytest>=7.0.0",
    "pytest-cov>=4.0.0",
    "black>=23.0.0",
    "flake8>=6.0.0",
    "mypy>=1.0.0",
]

[project.scripts]
ms-pg-replicator = "ms_pg_replicator.cli:main"

[tool.black]
line-length = 88
target-version = ['py38', 'py39', 'py310', 'py311', 'py312']

[tool.flake8]
max-line-length = 88
extend-ignore = ["E203"]

[tool.pytest.ini_options]
testpaths = ["tests"]
python_files = "test_*.py"
python_classes = "Test*"
python_functions = "test_*"
EOF

# Create requirements.txt (for backward compatibility)
cat > "$BASE_DIR/requirements.txt" << 'EOF'
# Core dependencies
pywin32>=311
psycopg2-binary>=2.9.0
pyyaml>=6.0

# Development dependencies (optional, for developers)
pytest>=7.0.0
pytest-cov>=4.0.0
black>=23.0.0
flake8>=6.0.0
mypy>=1.0.0
EOF

# Create README.md (starter template)
cat > "$BASE_DIR/README.md" << 'EOF'
# MS PG Replicator

A Python tool for replicating Microsoft Access databases to PostgreSQL.

## Description

Brief description of what your tool does and why it's useful.

## Installation

pip install ms-pg-replicator
EOF

