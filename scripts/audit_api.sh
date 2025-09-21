#!/bin/bash

# Audit API endpoints by grepping for FastAPI path operations
# Usage: ./scripts/audit_api.sh [backend/app/]

TARGET_DIR=${1:-backend/app/}

echo "=== Backend API Endpoints Audit ==="
echo "Scanning directory: $TARGET_DIR"
echo "Date: $(date)"
echo

# Find all Python files
find "$TARGET_DIR" -name "*.py" | while read -r file; do
    echo "File: $file"
    
    # Grep for @app. (direct in main.py)
    grep -n "^@app\." "$file" 2>/dev/null | while IFS=: read -r line_num line; do
        # Extract method and path
        if [[ $line =~ @app\.([a-z]+)\(\"?([^\" ]+) ]]; then
            method=${BASH_REMATCH[1]}
            path=${BASH_REMATCH[2]}
            echo "  Direct: $method $path"
        fi
    done
    
    # Grep for @router. (in router files)
    grep -n "^@router\." "$file" 2>/dev/null | while IFS=: read -r line_num line; do
        if [[ $line =~ @router\.([a-z]+)\(\"?([^\" ]+) ]]; then
            method=${BASH_REMATCH[1]}
            path=${BASH_REMATCH[2]}
            echo "  Router: $method $path"
        fi
    done
    
    echo
done

echo "=== End of Audit ==="
