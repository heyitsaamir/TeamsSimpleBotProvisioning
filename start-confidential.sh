#!/bin/bash

# Start script for Confidential Client version

echo "🔐 Starting Bot Provisioner (Confidential Client)..."
echo ""

# Check if CLIENT_SECRET is set
if [ -z "$CLIENT_SECRET" ]; then
    echo "❌ ERROR: CLIENT_SECRET environment variable not set"
    echo ""
    echo "Please set your Azure AD app client secret:"
    echo "  export CLIENT_SECRET=\"your-secret-value\""
    echo ""
    echo "Or run this script with:"
    echo "  CLIENT_SECRET=your-secret ./start-confidential.sh"
    echo ""
    exit 1
fi

echo "✓ Client secret found"
echo ""

# Start backend in background
echo "🚀 Starting backend server (port 3002)..."
cd backend
node server-confidential.js &
BACKEND_PID=$!
cd ..

# Wait for backend to start
sleep 2

# Start frontend
echo "🌐 Starting frontend server (port 8080)..."
cd frontend
python3 -m http.server 8080 &
FRONTEND_PID=$!
cd ..

echo ""
echo "✅ Both servers started!"
echo ""
echo "📋 Servers:"
echo "   Backend:  http://localhost:3002"
echo "   Frontend: http://localhost:8080"
echo ""
echo "🌐 Open in browser:"
echo "   http://localhost:8080/index-confidential.html"
echo ""
echo "Press Ctrl+C to stop both servers"
echo ""

# Trap Ctrl+C and kill both servers
trap "echo ''; echo '🛑 Stopping servers...'; kill $BACKEND_PID $FRONTEND_PID; exit 0" INT

# Wait for both processes
wait
