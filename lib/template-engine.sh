#!/bin/bash

# Template Engine for VBA Tool
# Processes templates and creates workbook structures

# Create workbook structure based on template
create_workbook_structure() {
    local project_path="$1"
    local workbook_name="$2"
    local template="$3"

    local workbook_path="$project_path/$workbook_name"
    local template_path="$TEMPLATES_DIR/$template"

    info "Creating workbook structure: $workbook_name (template: $template)"

    # Create workbook directory
    mkdir -p "$workbook_path"

    # Read template definition
    local template_json="$template_path/template.json"
    if [[ ! -f "$template_json" ]]; then
        error "Template definition not found: $template_json"
        exit 1
    fi

    # Create directories from template
    local folders=$(jq -r '.folders[]?' "$template_json")
    if [[ -n "$folders" ]]; then
        while IFS= read -r folder; do
            if [[ -n "$folder" ]]; then
                mkdir -p "$workbook_path/$folder"
                success "  Created: $folder/"
            fi
        done <<< "$folders"
    fi

    # Create LocalUtility.bas from template
    local template_localutil="$template_path/LocalUtility.bas"
    if [[ -f "$template_localutil" ]]; then
        # Replace placeholders
        local wsl_workbook_path="$project_path/$workbook_name"
        # Convert WSL path to Windows UNC path with proper separators
        local windows_workbook_path="\\\\wsl.localhost\\Ubuntu${wsl_workbook_path//\//\\}"
        # Escape backslashes for sed (prevents \U from being interpreted as uppercase directive)
        local windows_workbook_path_escaped=$(echo "$windows_workbook_path" | sed 's|\\|\\\\|g')

        sed -e "s|{{WSL_WORKBOOK_PATH}}|$wsl_workbook_path|g" \
            -e "s|{{WINDOWS_WORKBOOK_PATH}}|$windows_workbook_path_escaped|g" \
            -e "s|{{WORKBOOK_NAME}}|$workbook_name|g" \
            "$template_localutil" > "$workbook_path/LocalUtility.bas"

        success "  Created: LocalUtility.bas"
    else
        warning "  Template does not include LocalUtility.bas"
    fi

    # Create ProjectEntry.bas from template (version-controlled entry point)
    local template_projectentry="$template_path/ProjectEntry.bas"
    if [[ -f "$template_projectentry" ]]; then
        sed -e "s|{{WSL_WORKBOOK_PATH}}|$wsl_workbook_path|g" \
            -e "s|{{WINDOWS_WORKBOOK_PATH}}|$windows_workbook_path_escaped|g" \
            -e "s|{{WORKBOOK_NAME}}|$workbook_name|g" \
            "$template_projectentry" > "$workbook_path/ProjectEntry.bas"

        success "  Created: ProjectEntry.bas"
    else
        warning "  Template does not include ProjectEntry.bas"
    fi

    # Create placeholder files in folders if specified
    local placeholder_files=$(jq -r '.placeholder_files[]?' "$template_json")
    if [[ -n "$placeholder_files" ]]; then
        while IFS= read -r placeholder; do
            if [[ -n "$placeholder" ]]; then
                local placeholder_path="$workbook_path/$placeholder"
                local placeholder_dir=$(dirname "$placeholder_path")
                mkdir -p "$placeholder_dir"

                # Create .gitkeep to preserve empty directories
                touch "$placeholder_dir/.gitkeep"
            fi
        done <<< "$placeholder_files"
    fi

    success "Workbook structure created at: $workbook_path"
}

# Get WSL path for Windows path conversion
get_wsl_path() {
    local path="$1"
    # Convert /home/... to \\wsl.localhost\Ubuntu\home\...
    echo "\\\\wsl.localhost\\Ubuntu$path"
}

# Get Windows path for WSL path conversion
get_windows_path() {
    local path="$1"
    # Convert /mnt/c/... to C:\...
    echo "$path" | sed 's|/mnt/c/|C:\\|g; s|/|\\|g'
}
