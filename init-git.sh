#!/bin/bash

# Initialize git repository and set up remote
echo "🔧 Initializing Git repository..."

# Initialize git
git init

# Add remote
echo "🔗 Adding remote origin..."
git remote add origin https://github.com/heyitsaamir/TeamsSimpleBotProvisioning.git

# Rename branch to main
echo "🌿 Setting up main branch..."
git branch -M main

# Add all files
echo "📦 Adding files..."
git add .

# Create initial commit
echo "💾 Creating initial commit..."
git commit -m "Initial commit: Bot provisioner demo with MSAL authentication

Features:
- OAuth 2.0 Device Code Flow authentication
- MSAL integration with acquireTokenSilent()
- Single authentication for multiple resources (Graph + TDP)
- Azure AD app registration
- Teams app creation and bot registration
- Simple unstyled frontend wizard
- Express.js backend with session management
- Teams deep link for app installation

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>"

# Show status
echo "✅ Repository initialized!"
echo ""
echo "📊 Repository status:"
git status
echo ""
echo "📝 Commit log:"
git log --oneline -n 3
echo ""
echo "🚀 To push to GitHub, run:"
echo "   git push -u origin main"
