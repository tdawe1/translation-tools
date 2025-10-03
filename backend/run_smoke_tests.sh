#!/bin/bash
# Smoke test runner for the translation pipeline backend

set -e

echo "🔥 Running Smoke Tests for Translation Pipeline Backend"
echo "======================================================"

# Ensure we're in the backend directory
cd "$(dirname "$0")"

# Set environment for testing
export PYTEST_RUNNING=1
export DATABASE_URL=sqlite:///:memory:
export LOG_LEVEL=WARNING

# Create test directories if they don't exist
mkdir -p test_uploads test_outputs

# Function to run tests with specific markers
run_tests() {
    local marker=$1
    local description=$2

    echo ""
    echo "📋 Running $description..."
    echo "----------------------------------------"

    if [ -n "$marker" ]; then
        python -m pytest tests/ -v -m "$marker" --tb=short
    else
        python -m pytest tests/test_smoke_workflow.py -v --tb=short
    fi

    local exit_code=$?

    if [ $exit_code -eq 0 ]; then
        echo "✅ $description passed"
    else
        echo "❌ $description failed"
        return $exit_code
    fi
}

# Function to run specific test categories
run_category() {
    local category=$1
    echo ""
    echo "📋 Running $category tests..."
    echo "----------------------------------------"
    python -m pytest tests/test_smoke_workflow.py::Test$category -v --tb=short
}

# Main execution
echo ""
echo "🚀 Starting smoke test suite..."
echo ""

# Option 1: Run all smoke tests
if [ "$1" = "--all" ] || [ -z "$1" ]; then
    run_tests "" "all smoke tests"

    echo ""
    echo "📊 Running by category..."
    echo ""

    run_category "AuthenticationFlow"
    run_category "JobSubmissionWorkflow"
    run_category "ErrorScenarios"
    run_category "IntegrationPoints"

# Option 2: Run authentication tests only
elif [ "$1" = "--auth" ]; then
    run_category "AuthenticationFlow"

# Option 3: Run job workflow tests only
elif [ "$1" = "--jobs" ]; then
    run_category "JobSubmissionWorkflow"

# Option 4: Run error scenario tests only
elif [ "$1" = "--errors" ]; then
    run_category "ErrorScenarios"

# Option 5: Run integration tests only
elif [ "$1" = "--integration" ]; then
    run_category "IntegrationPoints"

# Option 6: Run with coverage
elif [ "$1" = "--coverage" ]; then
    echo "📊 Running tests with coverage report..."
    echo "----------------------------------------"
    python -m pytest tests/test_smoke_workflow.py --cov=app --cov-report=html --cov-report=term-missing

# Option 7: Run performance tests
elif [ "$1" = "--perf" ]; then
    run_tests "slow" "performance tests"

# Option 8: Run specific test file
elif [ "$1" = "--file" ] && [ -n "$2" ]; then
    echo "📋 Running tests from $2..."
    echo "----------------------------------------"
    python -m pytest "$2" -v --tb=short

else
    echo "Usage: $0 [OPTION]"
    echo ""
    echo "Options:"
    echo "  --all         Run all smoke tests (default)"
    echo "  --auth        Run authentication tests only"
    echo "  --jobs        Run job workflow tests only"
    echo "  --errors      Run error scenario tests only"
    echo "  --integration Run integration tests only"
    echo "  --coverage    Run tests with coverage report"
    echo "  --perf        Run performance tests"
    echo "  --file FILE   Run tests from specific file"
    echo ""
    echo "Examples:"
    echo "  $0                    # Run all smoke tests"
    echo "  $0 --auth             # Run authentication tests only"
    echo "  $0 --coverage         # Run with coverage report"
    exit 1
fi

# Cleanup
echo ""
echo "🧹 Cleaning up test artifacts..."
rm -rf test_uploads test_outputs __pycache__ .pytest_cache

echo ""
echo "✅ Smoke tests completed successfully!"
echo ""
echo "💡 Tip: Run './run_smoke_tests.sh --coverage' to see detailed coverage report"