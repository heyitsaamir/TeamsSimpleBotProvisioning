#!/bin/bash

# Bot Provisioner Demo - Quick Start Script

echo "🤖 Bot Provisioner Demo - Quick Start"
echo "====================================="
echo ""

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js 16+ first."
    echo "   Visit: https://nodejs.org/"
    exit 1
fi

echo "✅ Node.js version: $(node --version)"
echo ""

# Install backend dependencies if needed
if [ ! -d "backend/node_modules" ]; then
    echo "📦 Installing backend dependencies..."
    cd backend
    npm install
    cd ..
    echo ""
fi

# Start backend server
echo "🚀 Starting backend server on http://localhost:3001..."
cd backend
node server.js &
BACKEND_PID=$!
cd ..

# Wait for backend to start
sleep 2

# Check if backend is running
if curl -s http://localhost:3001/health > /dev/null; then
    echo "✅ Backend server is running!"
else
    echo "❌ Backend server failed to start"
    kill $BACKEND_PID 2>/dev/null
    exit 1
fi

echo ""
echo "🌐 Opening frontend in browser..."
echo ""

# Open frontend in browser
if command -v open &> /dev/null; then
    # macOS
    open "frontend/index.html"
elif command -v xdg-open &> /dev/null; then
    # Linux
    xdg-open "frontend/index.html"
elif command -v start &> /dev/null; then
    # Windows
    start "frontend/index.html"
fi

echo "✅ Demo is now running!"
echo ""
echo "📋 URLs:"
echo "   Backend:  http://localhost:3001"
echo "   Frontend: file://$(pwd)/frontend/index.html"
echo ""
echo "💡 Tip: For better CORS support, serve frontend with:"
echo "   cd frontend && python3 -m http.server 8080"
echo "   Then visit: http://localhost:8080"
echo ""
echo "Press Ctrl+C to stop the backend server"

# Wait for Ctrl+C
trap "echo ''; echo '🛑 Stopping backend server...'; kill $BACKEND_PID 2>/dev/null; exit 0" INT

# Keep script running
wait $BACKEND_PID
