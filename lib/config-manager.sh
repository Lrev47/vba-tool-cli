#!/bin/bash

# Configuration Manager for VBA Tool
# Handles loading and accessing configuration values

# Global config variables
WSL_BASE_PATH=""
WINDOWS_BASE_PATH=""
WINDOWS_BASE_PATH_WSL=""
DEFAULT_TEMPLATE=""
GIT_AUTO_INIT=""
DEFAULT_AUTHOR=""

# Load configuration from config.json
load_config() {
    if [[ ! -f "$CONFIG_FILE" ]]; then
        error "Configuration not found. Please run 'vba-tool init' first."
        exit 1
    fi

    # Check if jq is installed
    if ! command -v jq &> /dev/null; then
        error "jq is required but not installed. Please install: sudo apt install jq"
        exit 1
    fi

    # Load config values
    WSL_BASE_PATH=$(jq -r '.wsl_base_path' "$CONFIG_FILE")
    WINDOWS_BASE_PATH=$(jq -r '.windows_base_path' "$CONFIG_FILE")
    WINDOWS_BASE_PATH_WSL=$(jq -r '.windows_base_path_wsl' "$CONFIG_FILE")
    DEFAULT_TEMPLATE=$(jq -r '.default_template' "$CONFIG_FILE")
    GIT_AUTO_INIT=$(jq -r '.git_auto_init' "$CONFIG_FILE")
    DEFAULT_AUTHOR=$(jq -r '.default_author' "$CONFIG_FILE")

    # Validate required fields
    if [[ -z "$WSL_BASE_PATH" ]] || [[ "$WSL_BASE_PATH" == "null" ]]; then
        error "Invalid configuration: wsl_base_path is missing"
        exit 1
    fi
}

# Create .gitignore file
create_gitignore() {
    local project_path="$1"

    cat > "$project_path/.gitignore" << 'EOF'
# Local development utilities
**/LocalUtility.bas

# Excel temporary files
~$*

# Backup files
*.bak
*.tmp

# OS files
.DS_Store
Thumbs.db

# IDE files
.vscode/
.idea/

EOF
}

# Create project metadata
create_project_metadata() {
    local project_path="$1"
    local project_name="$2"
    local template="$3"
    local multi_workbook="$4"

    cat > "$project_path/.vba-project.json" << EOF
{
  "name": "$project_name",
  "template": "$template",
  "multi_workbook": $multi_workbook,
  "created": "$(date -Iseconds)",
  "author": "$DEFAULT_AUTHOR"
}
EOF
}
