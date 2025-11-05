"""
TouchPoint Tetris Game
Classic Tetris implementation using IronPython 2.7.3 and HTML5 Canvas
Access via: /PyScriptForm/TetrisGame
"""

import json
import datetime

# Handle high score submission
if model.HttpMethod == "post" and hasattr(Data, 'action'):
    action = Data.action

    if action == 'submit_score':
        # Load existing high scores
        try:
            scores_json = model.TextContent("TetrisHighScores")
            scores = json.loads(scores_json) if scores_json else []
        except:
            scores = []

        # Add new score
        new_score = {
            'player': Data.player if hasattr(Data, 'player') else 'Anonymous',
            'score': int(Data.score),
            'level': int(Data.level),
            'lines': int(Data.lines),
            'date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        }

        scores.append(new_score)

        # Sort by score (descending) and keep top 10
        scores = sorted(scores, key=lambda x: x['score'], reverse=True)[:10]

        # Save back to content storage
        model.WriteContentText("TetrisHighScores", json.dumps(scores), "")

        # Return updated leaderboard
        print json.dumps({'success': True, 'scores': scores})
        return

    elif action == 'get_scores':
        # Return current high scores
        try:
            scores_json = model.TextContent("TetrisHighScores")
            scores = json.loads(scores_json) if scores_json else []
        except:
            scores = []

        print json.dumps({'success': True, 'scores': scores})
        return

# Generate the game HTML
model.Form = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>TouchPoint Tetris</title>
    <style>
        body {
            margin: 0;
            padding: 20px;
            font-family: 'Courier New', monospace;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }

        .game-container {
            display: flex;
            gap: 30px;
            background: rgba(0, 0, 0, 0.3);
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.5);
        }

        .game-area {
            display: flex;
            flex-direction: column;
            align-items: center;
        }

        canvas {
            border: 3px solid #fff;
            box-shadow: 0 0 20px rgba(255, 255, 255, 0.3);
            background: #000;
        }

        .controls {
            margin-top: 20px;
            text-align: center;
        }

        .controls h3 {
            margin: 10px 0;
            font-size: 16px;
        }

        .key {
            background: rgba(255, 255, 255, 0.2);
            padding: 5px 10px;
            border-radius: 5px;
            display: inline-block;
            margin: 2px;
            font-weight: bold;
        }

        .sidebar {
            display: flex;
            flex-direction: column;
            gap: 20px;
        }

        .info-panel {
            background: rgba(255, 255, 255, 0.1);
            padding: 20px;
            border-radius: 10px;
            min-width: 200px;
        }

        .info-panel h2 {
            margin: 0 0 15px 0;
            font-size: 20px;
            border-bottom: 2px solid rgba(255, 255, 255, 0.3);
            padding-bottom: 10px;
        }

        .stat {
            display: flex;
            justify-content: space-between;
            margin: 10px 0;
            font-size: 16px;
        }

        .stat-label {
            font-weight: bold;
        }

        .stat-value {
            background: rgba(255, 255, 255, 0.2);
            padding: 2px 10px;
            border-radius: 5px;
        }

        .next-piece {
            text-align: center;
        }

        #nextCanvas {
            border: 2px solid rgba(255, 255, 255, 0.5);
            margin-top: 10px;
            background: #000;
        }

        button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: 2px solid white;
            padding: 10px 20px;
            font-size: 16px;
            font-weight: bold;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s;
            font-family: 'Courier New', monospace;
        }

        button:hover {
            transform: scale(1.05);
            box-shadow: 0 5px 15px rgba(255, 255, 255, 0.3);
        }

        button:active {
            transform: scale(0.95);
        }

        .leaderboard {
            max-height: 400px;
            overflow-y: auto;
        }

        .leaderboard table {
            width: 100%;
            border-collapse: collapse;
        }

        .leaderboard th {
            background: rgba(255, 255, 255, 0.2);
            padding: 8px;
            text-align: left;
            font-size: 14px;
        }

        .leaderboard td {
            padding: 6px 8px;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
            font-size: 13px;
        }

        .rank-1 { color: #FFD700; font-weight: bold; }
        .rank-2 { color: #C0C0C0; font-weight: bold; }
        .rank-3 { color: #CD7F32; font-weight: bold; }

        .game-over-modal {
            display: none;
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: rgba(0, 0, 0, 0.95);
            padding: 40px;
            border-radius: 15px;
            border: 3px solid white;
            text-align: center;
            z-index: 1000;
        }

        .game-over-modal.show {
            display: block;
        }

        .game-over-modal h1 {
            color: #ff4444;
            margin: 0 0 20px 0;
            font-size: 48px;
        }

        .game-over-modal input {
            background: rgba(255, 255, 255, 0.1);
            border: 2px solid white;
            color: white;
            padding: 10px;
            font-size: 16px;
            border-radius: 5px;
            margin: 10px 0;
            width: 200px;
            font-family: 'Courier New', monospace;
        }

        .game-over-modal input::placeholder {
            color: rgba(255, 255, 255, 0.5);
        }

        #pauseOverlay {
            display: none;
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            justify-content: center;
            align-items: center;
            font-size: 48px;
            font-weight: bold;
            z-index: 100;
        }

        #pauseOverlay.show {
            display: flex;
        }
    </style>
</head>
<body>
    <div class="game-container">
        <div class="game-area">
            <h1 style="margin: 0 0 20px 0; text-align: center;">🎮 TETRIS 🎮</h1>
            <div style="position: relative;">
                <canvas id="gameCanvas" width="300" height="600"></canvas>
                <div id="pauseOverlay">PAUSED</div>
            </div>
            <div class="controls">
                <button id="startBtn" onclick="startGame()">START GAME</button>
                <button id="pauseBtn" onclick="togglePause()" style="display:none;">PAUSE</button>
                <h3 style="margin-top: 20px;">Controls:</h3>
                <div>
                    <span class="key">←</span> Move Left
                    <span class="key">→</span> Move Right
                    <span class="key">↓</span> Soft Drop
                </div>
                <div>
                    <span class="key">↑</span> Rotate
                    <span class="key">Space</span> Hard Drop
                    <span class="key">P</span> Pause
                </div>
            </div>
        </div>

        <div class="sidebar">
            <div class="info-panel">
                <h2>Stats</h2>
                <div class="stat">
                    <span class="stat-label">Score:</span>
                    <span class="stat-value" id="score">0</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Level:</span>
                    <span class="stat-value" id="level">1</span>
                </div>
                <div class="stat">
                    <span class="stat-label">Lines:</span>
                    <span class="stat-value" id="lines">0</span>
                </div>
            </div>

            <div class="info-panel next-piece">
                <h2>Next Piece</h2>
                <canvas id="nextCanvas" width="120" height="120"></canvas>
            </div>

            <div class="info-panel leaderboard">
                <h2>High Scores</h2>
                <table id="leaderboardTable">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Player</th>
                            <th>Score</th>
                        </tr>
                    </thead>
                    <tbody id="leaderboardBody">
                        <tr><td colspan="3">Loading...</td></tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div id="gameOverModal" class="game-over-modal">
        <h1>GAME OVER</h1>
        <div class="stat">
            <span class="stat-label">Final Score:</span>
            <span class="stat-value" id="finalScore">0</span>
        </div>
        <div class="stat">
            <span class="stat-label">Level Reached:</span>
            <span class="stat-value" id="finalLevel">1</span>
        </div>
        <div class="stat">
            <span class="stat-label">Lines Cleared:</span>
            <span class="stat-value" id="finalLines">0</span>
        </div>
        <div style="margin-top: 30px;">
            <input type="text" id="playerName" placeholder="Enter your name" maxlength="20" />
            <br/>
            <button onclick="submitScore()">SUBMIT SCORE</button>
            <button onclick="restartGame()">PLAY AGAIN</button>
        </div>
    </div>

    <script>
        // Game constants
        const COLS = 10;
        const ROWS = 20;
        const BLOCK_SIZE = 30;
        const COLORS = [
            null,
            '#FF0D72', // I
            '#0DC2FF', // O
            '#0DFF72', // T
            '#F538FF', // S
            '#FF8E0D', // Z
            '#FFE138', // J
            '#3877FF'  // L
        ];

        // Tetromino shapes
        const SHAPES = [
            [],
            [[0,0,0,0], [1,1,1,1], [0,0,0,0], [0,0,0,0]], // I
            [[2,2], [2,2]], // O
            [[0,3,0], [3,3,3], [0,0,0]], // T
            [[0,4,4], [4,4,0], [0,0,0]], // S
            [[5,5,0], [0,5,5], [0,0,0]], // Z
            [[6,0,0], [6,6,6], [0,0,0]], // J
            [[0,0,7], [7,7,7], [0,0,0]]  // L
        ];

        // Game state
        let canvas = document.getElementById('gameCanvas');
        let ctx = canvas.getContext('2d');
        let nextCanvas = document.getElementById('nextCanvas');
        let nextCtx = nextCanvas.getContext('2d');

        let board = [];
        let currentPiece = null;
        let nextPiece = null;
        let score = 0;
        let level = 1;
        let lines = 0;
        let gameRunning = false;
        let gamePaused = false;
        let dropCounter = 0;
        let dropInterval = 1000;
        let lastTime = 0;

        // Initialize board
        function createBoard() {
            board = Array(ROWS).fill().map(() => Array(COLS).fill(0));
        }

        // Create new piece
        function createPiece(type) {
            return {
                shape: SHAPES[type],
                type: type,
                x: Math.floor(COLS / 2) - Math.floor(SHAPES[type][0].length / 2),
                y: 0
            };
        }

        function randomPiece() {
            return createPiece(Math.floor(Math.random() * 7) + 1);
        }

        // Draw functions
        function drawBlock(ctx, x, y, type) {
            ctx.fillStyle = COLORS[type];
            ctx.fillRect(x * BLOCK_SIZE, y * BLOCK_SIZE, BLOCK_SIZE, BLOCK_SIZE);
            ctx.strokeStyle = '#000';
            ctx.lineWidth = 2;
            ctx.strokeRect(x * BLOCK_SIZE, y * BLOCK_SIZE, BLOCK_SIZE, BLOCK_SIZE);
        }

        function drawBoard() {
            ctx.fillStyle = '#000';
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            for (let y = 0; y < ROWS; y++) {
                for (let x = 0; x < COLS; x++) {
                    if (board[y][x]) {
                        drawBlock(ctx, x, y, board[y][x]);
                    }
                }
            }
        }

        function drawPiece(piece, context, offsetX = 0, offsetY = 0) {
            piece.shape.forEach((row, y) => {
                row.forEach((value, x) => {
                    if (value) {
                        drawBlock(context, piece.x + x + offsetX, piece.y + y + offsetY, piece.type);
                    }
                });
            });
        }

        function drawNextPiece() {
            nextCtx.fillStyle = '#000';
            nextCtx.fillRect(0, 0, nextCanvas.width, nextCanvas.height);

            if (nextPiece) {
                let offsetX = (4 - nextPiece.shape[0].length) / 2;
                let offsetY = (4 - nextPiece.shape.length) / 2;

                nextPiece.shape.forEach((row, y) => {
                    row.forEach((value, x) => {
                        if (value) {
                            drawBlock(nextCtx, x + offsetX, y + offsetY, nextPiece.type);
                        }
                    });
                });
            }
        }

        function draw() {
            drawBoard();
            if (currentPiece) {
                drawPiece(currentPiece, ctx);
            }
            drawNextPiece();
        }

        // Collision detection
        function collides(piece, offsetX = 0, offsetY = 0) {
            for (let y = 0; y < piece.shape.length; y++) {
                for (let x = 0; x < piece.shape[y].length; x++) {
                    if (piece.shape[y][x]) {
                        let newX = piece.x + x + offsetX;
                        let newY = piece.y + y + offsetY;

                        if (newX < 0 || newX >= COLS || newY >= ROWS) {
                            return true;
                        }

                        if (newY >= 0 && board[newY][newX]) {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        // Piece movement
        function movePiece(dir) {
            currentPiece.x += dir;
            if (collides(currentPiece)) {
                currentPiece.x -= dir;
                return false;
            }
            return true;
        }

        function rotatePiece() {
            let rotated = currentPiece.shape[0].map((_, i) =>
                currentPiece.shape.map(row => row[i]).reverse()
            );

            let previousShape = currentPiece.shape;
            currentPiece.shape = rotated;

            // Wall kick
            let offset = 0;
            while (collides(currentPiece, offset)) {
                offset = offset > 0 ? -(offset + 1) : -offset;
                if (Math.abs(offset) > currentPiece.shape[0].length) {
                    currentPiece.shape = previousShape;
                    return;
                }
            }

            currentPiece.x += offset;
        }

        function dropPiece() {
            currentPiece.y++;
            if (collides(currentPiece)) {
                currentPiece.y--;
                mergePiece();
                clearLines();
                nextPieceToBoard();

                if (collides(currentPiece)) {
                    gameOver();
                }
                return false;
            }
            return true;
        }

        function hardDrop() {
            while (dropPiece()) {
                score += 2;
            }
            updateStats();
        }

        function mergePiece() {
            currentPiece.shape.forEach((row, y) => {
                row.forEach((value, x) => {
                    if (value) {
                        board[currentPiece.y + y][currentPiece.x + x] = currentPiece.type;
                    }
                });
            });
        }

        function clearLines() {
            let linesCleared = 0;

            outer: for (let y = ROWS - 1; y >= 0; y--) {
                for (let x = 0; x < COLS; x++) {
                    if (!board[y][x]) {
                        continue outer;
                    }
                }

                // Remove line
                board.splice(y, 1);
                board.unshift(Array(COLS).fill(0));
                linesCleared++;
                y++;
            }

            if (linesCleared > 0) {
                lines += linesCleared;

                // Scoring
                const lineScores = [0, 100, 300, 500, 800];
                score += lineScores[linesCleared] * level;

                // Level up every 10 lines
                level = Math.floor(lines / 10) + 1;
                dropInterval = Math.max(100, 1000 - (level - 1) * 100);

                updateStats();
            }
        }

        function nextPieceToBoard() {
            if (!nextPiece) {
                nextPiece = randomPiece();
            }
            currentPiece = nextPiece;
            nextPiece = randomPiece();
        }

        function updateStats() {
            document.getElementById('score').textContent = score;
            document.getElementById('level').textContent = level;
            document.getElementById('lines').textContent = lines;
        }

        // Game loop
        function gameLoop(time = 0) {
            if (!gameRunning || gamePaused) {
                if (gameRunning) {
                    requestAnimationFrame(gameLoop);
                }
                return;
            }

            const deltaTime = time - lastTime;
            lastTime = time;

            dropCounter += deltaTime;
            if (dropCounter > dropInterval) {
                dropPiece();
                dropCounter = 0;
            }

            draw();
            requestAnimationFrame(gameLoop);
        }

        // Game controls
        function startGame() {
            createBoard();
            score = 0;
            level = 1;
            lines = 0;
            dropCounter = 0;
            dropInterval = 1000;

            nextPiece = randomPiece();
            nextPieceToBoard();

            gameRunning = true;
            gamePaused = false;

            document.getElementById('startBtn').style.display = 'none';
            document.getElementById('pauseBtn').style.display = 'inline-block';
            document.getElementById('gameOverModal').classList.remove('show');

            updateStats();
            requestAnimationFrame(gameLoop);
        }

        function togglePause() {
            gamePaused = !gamePaused;
            document.getElementById('pauseOverlay').classList.toggle('show');
            document.getElementById('pauseBtn').textContent = gamePaused ? 'RESUME' : 'PAUSE';
        }

        function gameOver() {
            gameRunning = false;

            document.getElementById('finalScore').textContent = score;
            document.getElementById('finalLevel').textContent = level;
            document.getElementById('finalLines').textContent = lines;
            document.getElementById('gameOverModal').classList.add('show');
            document.getElementById('playerName').value = '';
            document.getElementById('playerName').focus();

            document.getElementById('startBtn').style.display = 'inline-block';
            document.getElementById('pauseBtn').style.display = 'none';
        }

        function restartGame() {
            document.getElementById('gameOverModal').classList.remove('show');
            startGame();
        }

        function submitScore() {
            let playerName = document.getElementById('playerName').value.trim() || 'Anonymous';

            // Get script name from URL
            var pathname = window.location.pathname;
            var scriptName = pathname.split('/').pop().split('?')[0];

            fetch('/PyScriptForm/' + scriptName, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: 'action=submit_score&player=' + encodeURIComponent(playerName) +
                      '&score=' + score + '&level=' + level + '&lines=' + lines
            })
            .then(response => response.json())
            .then(data => {
                loadLeaderboard();
                document.getElementById('gameOverModal').classList.remove('show');
            })
            .catch(error => console.error('Error:', error));
        }

        function loadLeaderboard() {
            // Get script name from URL
            var pathname = window.location.pathname;
            var scriptName = pathname.split('/').pop().split('?')[0];

            fetch('/PyScriptForm/' + scriptName + '?action=get_scores')
            .then(response => response.json())
            .then(data => {
                let tbody = document.getElementById('leaderboardBody');
                tbody.innerHTML = '';

                if (data.scores && data.scores.length > 0) {
                    data.scores.forEach((entry, index) => {
                        let tr = document.createElement('tr');
                        tr.className = 'rank-' + (index + 1);
                        tr.innerHTML =
                            '<td>' + (index + 1) + '</td>' +
                            '<td>' + entry.player + '</td>' +
                            '<td>' + entry.score + '</td>';
                        tbody.appendChild(tr);
                    });
                } else {
                    tbody.innerHTML = '<tr><td colspan="3">No scores yet!</td></tr>';
                }
            })
            .catch(error => {
                console.error('Error:', error);
                document.getElementById('leaderboardBody').innerHTML =
                    '<tr><td colspan="3">Error loading scores</td></tr>';
            });
        }

        // Keyboard controls
        document.addEventListener('keydown', (e) => {
            if (!gameRunning || gamePaused) {
                if (e.key === 'p' || e.key === 'P') {
                    if (gameRunning) togglePause();
                }
                return;
            }

            switch(e.key) {
                case 'ArrowLeft':
                    movePiece(-1);
                    break;
                case 'ArrowRight':
                    movePiece(1);
                    break;
                case 'ArrowDown':
                    if (dropPiece()) {
                        score += 1;
                        updateStats();
                    }
                    break;
                case 'ArrowUp':
                    rotatePiece();
                    break;
                case ' ':
                    e.preventDefault();
                    hardDrop();
                    break;
                case 'p':
                case 'P':
                    togglePause();
                    break;
            }

            draw();
        });

        // Allow Enter to submit score
        document.getElementById('playerName').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                submitScore();
            }
        });

        // Load leaderboard on page load
        loadLeaderboard();

        // Initial draw
        createBoard();
        draw();
    </script>
</body>
</html>
'''
