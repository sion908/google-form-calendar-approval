#!/bin/bash

# Load environment variables
if [ -f .env ]; then
    export $(grep -v '^#' .env | xargs)
else
    echo ".env file not found. Please create one."
    exit 1
fi

# Function to show help
show_help() {
    echo "Usage: $0 [command] [options]"
    echo ""
    echo "Commands:"
    echo "  deploy      Deploy the application"
    echo "  push        Push changes to the repository"
    echo "  commit      Commit changes with a message"
    echo "  full        Run full deployment (commit, push, deploy)"
    echo "  help        Show this help message"
    echo ""
    echo "Options:"
    echo "  -m, --message   Commit message (for commit command)"
    echo "  -i, --id        Deployment ID (overrides .env)"
}

# Function to commit changes
commit_changes() {
    local message="$1"
    if [ -z "$message" ]; then
        message="$GIT_COMMIT_MESSAGE"
    fi
    
    echo "Staging changes..."
    git add .
    
    echo "Committing changes with message: $message"
    git commit -m "$message"
}

# Function to push changes
push_changes() {
    echo "Pushing changes to repository..."
    git push
}

# Function to deploy the app
deploy_app() {
    local deploy_id="$1"
    if [ -z "$deploy_id" ]; then
        deploy_id="$DEPLOYMENT_ID"
    fi
    
    if [ -z "$deploy_id" ]; then
        echo "Error: Deployment ID not provided. Please set DEPLOYMENT_ID in .env or use -i option."
        exit 1
    fi
    
    echo "Deploying with ID: $deploy_id"
    clasp deploy -i "$deploy_id"
}

# Main script
case "$1" in
    commit)
        shift
        message=""
        while [[ $# -gt 0 ]]; do
            case "$1" in
                -m|--message)
                    message="$2"
                    shift 2
                    ;;
                *)
                    echo "Unknown option: $1"
                    show_help
                    exit 1
                    ;;
            esac
        done
        commit_changes "$message"
        ;;
        
    push)
        push_changes
        ;;
        
    deploy)
        shift
        deploy_id=""
        while [[ $# -gt 0 ]]; do
            case "$1" in
                -i|--id)
                    deploy_id="$2"
                    shift 2
                    ;;
                *)
                    echo "Unknown option: $1"
                    show_help
                    exit 1
                    ;;
            esac
        done
        deploy_app "$deploy_id"
        ;;
        
    full)
        shift
        message=""
        deploy_id=""
        
        # Process all arguments
        while [[ $# -gt 0 ]]; do
            case "$1" in
                -m|--message)
                    message="$2"
                    shift 2
                    ;;
                -i|--id)
                    deploy_id="$2"
                    shift 2
                    ;;
                *)
                    echo "Unknown option: $1"
                    show_help
                    exit 1
                    ;;
            esac
        done
        
        commit_changes "$message"
        push_changes
        deploy_app "$deploy_id"
        ;;
        
    help|--help|-h)
        show_help
        ;;
        
    *)
        echo "Unknown command: $1"
        show_help
        exit 1
        ;;
esac

exit 0
