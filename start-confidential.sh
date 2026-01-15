#!/bin/bash

# Bot Provisioner Demo - Confidential Client Start Script

echo "🤖 Bot Provisioner Demo - Confidential Client"
echo "================================================"
echo ""

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js 16+ first."
    echo "   Visit: https://nodejs.org/"
    exit 1
fi

echo "✅ Node.js version: $(node --version)"
echo ""

# Check if Python is installed for serving frontend
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 is not installed. Please install Python3 first."
    exit 1
fi

echo "✅ Python3 version: $(python3 --version)"
echo ""

# Check if CLIENT_SECRET is set
if [ -z "$CLIENT_SECRET" ]; then
    echo "⚠️  CLIENT_SECRET environment variable is not set"
    echo "   Set it with: export CLIENT_SECRET=your_secret_here"
    echo "   Continuing anyway (will use placeholder)..."
    echo ""
fi

# Install backend dependencies if needed
if [ ! -d "backend/node_modules" ]; then
    echo "📦 Installing backend dependencies..."
    cd backend
    npm install
    cd ..
    echo ""
fi

# Start backend server
echo "🚀 Starting backend server on http://localhost:3003..."
cd backend
node server-confidential.js &
BACKEND_PID=$!
cd ..

# Wait for backend to start
sleep 2

# Check if backend is running
if curl -s http://localhost:3003/health > /dev/null; then
    echo "✅ Backend server is running!"
else
    echo "❌ Backend server failed to start"
    kill $BACKEND_PID 2>/dev/null
    exit 1
fi

echo ""

# Start frontend server
echo "🌐 Starting frontend server on http://localhost:8080..."
cd frontend
python3 -m http.server 8080 > /dev/null 2>&1 &
FRONTEND_PID=$!
cd ..

# Wait for frontend to start
sleep 2

echo "✅ Frontend server is running!"
echo ""

# Open frontend in browser
echo "🌐 Opening frontend in browser..."
if command -v open &> /dev/null; then
    # macOS
    open "http://localhost:8080/index-confidential.html"
elif command -v xdg-open &> /dev/null; then
    # Linux
    xdg-open "http://localhost:8080/index-confidential.html"
elif command -v start &> /dev/null; then
    # Windows
    start "http://localhost:8080/index-confidential.html"
fi

echo ""
echo "✅ Demo is now running!"
echo ""
echo "📋 URLs:"
echo "   Backend:  http://localhost:3003"
echo "   Frontend: http://localhost:8080/index-confidential.html"
echo ""
echo "📝 Next steps:"
echo "   1. Make sure your devtunnel is running"
echo "   2. Update redirect URIs in Azure AD:"
echo "      - https://YOUR_TUNNEL/redirect.html"
echo "      - https://YOUR_TUNNEL/admin-consent-callback.html"
echo "   3. Open: https://YOUR_TUNNEL/index-confidential.html"
echo ""
echo "Press Ctrl+C to stop all servers"

# Wait for Ctrl+C
trap "echo ''; echo '🛑 Stopping servers...'; kill $BACKEND_PID $FRONTEND_PID 2>/dev/null; exit 0" INT

# Keep script running
wait $BACKEND_PID $FRONTEND_PID
