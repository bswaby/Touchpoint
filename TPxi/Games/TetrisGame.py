#Roles=Access
# Tetris Game
# Created By: Ben Swaby
# Email: bswaby@fbchtn.org

import json
import random
import datetime

model.Header = 'Tetris Game'

# Tetris pieces (tetrominoes) - represented as 4x4 grids
PIECES = {
    'I': [[0,0,0,0], [1,1,1,1], [0,0,0,0], [0,0,0,0]],
    'O': [[1,1], [1,1]],
    'T': [[0,1,0], [1,1,1], [0,0,0]],
    'S': [[0,1,1], [1,1,0], [0,0,0]],
    'Z': [[1,1,0], [0,1,1], [0,0,0]],
    'J': [[1,0,0], [1,1,1], [0,0,0]],
    'L': [[0,0,1], [1,1,1], [0,0,0]]
}

PIECE_COLORS = {
    'I': '#00f0f0',  # Cyan
    'O': '#f0f000',  # Yellow
    'T': '#a000f0',  # Purple
    'S': '#00f000',  # Green
    'Z': '#f00000',  # Red
    'J': '#0000f0',  # Blue
    'L': '#f0a000'   # Orange
}

# Helper functions (must be defined before AJAX handlers)
def rotate_piece(piece, rotation):
    """Rotate piece matrix"""
    result = piece
    for _ in range(rotation % 4):
        n = len(result)
        rotated = [[0 for _ in range(n)] for _ in range(n)]
        for i in range(n):
            for j in range(n):
                rotated[i][j] = result[n - 1 - j][i]
        result = rotated
    return result

def is_valid_position(board, piece, x, y, rotation):
    """Check if piece position is valid"""
    rotated = rotate_piece(piece, rotation)

    for py in range(len(rotated)):
        for px in range(len(rotated[py])):
            if rotated[py][px]:
                board_x = x + px
                board_y = y + py

                # Check boundaries
                if board_x < 0 or board_x >= 10 or board_y >= 20:
                    return False

                # Check collision with locked blocks
                if board_y >= 0 and board[board_y][board_x]:
                    return False

    return True

def lock_piece(game_state):
    """Lock current piece to board"""
    piece = game_state['currentPiece']
    piece_shape = PIECES[piece['type']]
    rotated = rotate_piece(piece_shape, piece['rotation'])
    color = PIECE_COLORS[piece['type']]

    for y in range(len(rotated)):
        for x in range(len(rotated[y])):
            if rotated[y][x]:
                board_y = piece['y'] + y
                board_x = piece['x'] + x
                if 0 <= board_y < 20 and 0 <= board_x < 10:
                    game_state['board'][board_y][board_x] = color

def clear_lines(game_state):
    """Clear completed lines and return count"""
    lines_cleared = 0
    y = 19

    while y >= 0:
        if all(game_state['board'][y]):
            # Remove line
            del game_state['board'][y]
            # Add new empty line at top
            game_state['board'].insert(0, [0 for _ in range(10)])
            lines_cleared += 1
        else:
            y -= 1

    return lines_cleared

# Handle AJAX requests
if model.HttpMethod == "post" and hasattr(Data, 'action'):
    action = Data.action

    if action == 'init':
        # Initialize new game
        piece_types = ['I', 'O', 'T', 'S', 'Z', 'J', 'L']
        current_type = random.choice(piece_types)
        next_type = random.choice(piece_types)

        game_state = {
            'board': [[0 for _ in range(10)] for _ in range(20)],  # 10x20 grid
            'currentPiece': {
                'type': current_type,
                'x': 3,
                'y': 0,
                'rotation': 0
            },
            'nextPiece': next_type,
            'score': 0,
            'lines': 0,
            'level': 1,
            'gameOver': False,
            'highScore': 0
        }

        # Get user's best score from leaderboard
        try:
            person = model.GetPerson(model.UserPeopleId)
            player_name = person.Name if person else 'Unknown Player'

            scores_json = model.TextContent('TetrisHighScores')
            all_scores = json.loads(scores_json) if scores_json else []

            user_best = 0
            for score_entry in all_scores:
                if score_entry.get('name') == player_name:
                    user_best = max(user_best, score_entry.get('score', 0))

            game_state['highScore'] = user_best
            game_state['playerName'] = player_name
        except:
            game_state['playerName'] = 'Unknown Player'
            pass

        print json.dumps({'success': True, 'gameState': game_state, 'pieces': PIECES, 'colors': PIECE_COLORS})

    elif action == 'saveHighScore':
        # Save high score to shared leaderboard
        score = int(Data.score) if hasattr(Data, 'score') else 0

        try:
            # Get current user's name
            person = model.GetPerson(model.UserPeopleId)
            player_name = person.Name if person else 'Unknown Player'

            # Load existing scores
            try:
                scores_json = model.TextContent('TetrisHighScores')
                scores = json.loads(scores_json) if scores_json else []
            except:
                scores = []

            # Add new score with timestamp
            from datetime import datetime
            scores.append({
                'name': player_name,
                'score': score,
                'date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })

            # Sort by score (descending) and keep top 50
            scores.sort(key=lambda x: x['score'], reverse=True)
            scores = scores[:50]

            # Save back to content
            model.WriteContentText('TetrisHighScores', json.dumps(scores), '')

            print json.dumps({'success': True, 'playerName': player_name})
        except Exception as e:
            print json.dumps({'success': False, 'error': str(e)})

    elif action == 'getHighScores':
        # Get top 10 high scores from content storage
        try:
            scores_json = model.TextContent('TetrisHighScores')
            all_scores = json.loads(scores_json) if scores_json else []

            # Return top 10
            top_scores = all_scores[:10]

            # Get current user's best score
            person = model.GetPerson(model.UserPeopleId)
            player_name = person.Name if person else 'Unknown Player'

            user_best = 0
            for score_entry in all_scores:
                if score_entry.get('name') == player_name:
                    user_best = max(user_best, score_entry.get('score', 0))
                    break

            print json.dumps({
                'success': True,
                'scores': top_scores,
                'userBest': user_best,
                'playerName': player_name
            })
        except Exception as e:
            print json.dumps({'success': True, 'scores': [], 'userBest': 0, 'error': str(e)})

else:
    # Display the game interface
    model.Form = '''
<!DOCTYPE html>
<html>
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tetris Game</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        html, body {
            width: 100%;
            height: 100%;
            overflow: auto;
        }

        body {
            font-family: 'Arial', sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            padding: 20px;
        }

        .game-container {
            display: flex;
            gap: 30px;
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            padding: 30px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.4);
            margin: 0 auto;
            max-width: 1000px;
            width: fit-content;
        }

        .main-panel {
            display: flex;
            flex-direction: column;
            align-items: center;
        }

        h1 {
            color: #1e3c72;
            margin-bottom: 20px;
            font-size: 36px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }

        #game-board {
            width: 300px;
            height: 600px;
            border: 3px solid #333;
            position: relative;
            background: #000;
            box-shadow: inset 0 0 30px rgba(0,0,0,0.8);
            border-radius: 5px;
        }

        canvas {
            display: block;
            width: 100%;
            height: 100%;
        }

        .side-panel {
            display: flex;
            flex-direction: column;
            gap: 20px;
            min-width: 200px;
        }

        .panel-section {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 15px;
            border-radius: 10px;
            color: white;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        }

        .panel-section h3 {
            margin-bottom: 10px;
            font-size: 18px;
        }

        .stat {
            display: flex;
            justify-content: space-between;
            margin: 5px 0;
            font-size: 16px;
        }

        .stat-value {
            font-weight: bold;
        }

        #next-piece {
            width: 100px;
            height: 100px;
            background: #000;
            margin: 10px auto;
            border: 2px solid #333;
            border-radius: 5px;
        }

        .controls {
            margin-top: 20px;
            text-align: center;
        }

        button {
            padding: 12px 24px;
            font-size: 16px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 25px;
            cursor: pointer;
            margin: 5px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            transition: all 0.3s;
        }

        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0,0,0,0.3);
        }

        button:active {
            transform: translateY(0);
        }

        .game-over {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: rgba(0, 0, 0, 0.95);
            color: white;
            padding: 40px;
            border-radius: 15px;
            text-align: center;
            display: none;
            z-index: 1000;
            box-shadow: 0 10px 40px rgba(0,0,0,0.5);
        }

        .game-over.show {
            display: block;
            animation: fadeIn 0.3s;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translate(-50%, -50%) scale(0.8); }
            to { opacity: 1; transform: translate(-50%, -50%) scale(1); }
        }

        .game-over h2 {
            margin-bottom: 20px;
            font-size: 32px;
        }

        .instructions {
            margin-top: 20px;
            color: #666;
            font-size: 14px;
            text-align: center;
        }

        .key-hint {
            display: inline-block;
            padding: 5px 10px;
            background: #f0f0f0;
            border: 2px solid #ccc;
            border-radius: 5px;
            margin: 2px;
            font-weight: bold;
        }

        .high-scores {
            max-height: 300px;
            overflow-y: auto;
        }

        .high-scores ol {
            list-style-position: inside;
        }

        .high-scores li {
            margin: 5px 0;
            font-size: 14px;
        }

        @media (max-width: 900px) {
            .game-container {
                flex-direction: column;
                padding: 20px;
                width: auto;
            }

            .side-panel {
                width: 100%;
            }
        }

        @media (max-width: 600px) {
            body {
                padding: 10px;
            }

            .game-container {
                padding: 15px;
            }

            #game-board {
                width: 250px;
                height: 500px;
            }

            h1 {
                font-size: 28px;
            }
        }
    </style>
</head>
<body>
    <div class="game-container">
        <div class="main-panel">
            <h1>🎮 Tetris</h1>

            <div id="game-board">
                <canvas id="canvas" width="300" height="600"></canvas>
                <div class="game-over" id="gameOver">
                    <h2>Game Over!</h2>
                    <p style="font-size: 20px; margin: 10px 0;">Score: <span id="finalScore">0</span></p>
                    <p style="font-size: 16px; margin: 10px 0;">Lines: <span id="finalLines">0</span></p>
                    <button onclick="startNewGame()" style="margin-top: 20px;">Play Again</button>
                </div>
            </div>

            <div class="controls">
                <button onclick="startNewGame()">New Game</button>
                <button onclick="togglePause()">Pause</button>
            </div>

            <div class="instructions">
                <p><span class="key-hint">←</span> <span class="key-hint">→</span> Move |
                   <span class="key-hint">↑</span> Rotate |
                   <span class="key-hint">↓</span> Soft Drop |
                   <span class="key-hint">SPACE</span> Hard Drop</p>
            </div>
        </div>

        <div class="side-panel">
            <div class="panel-section">
                <h3>📊 Stats</h3>
                <div class="stat">
                    <span>Score:</span>
                    <span class="stat-value" id="score">0</span>
                </div>
                <div class="stat">
                    <span>Lines:</span>
                    <span class="stat-value" id="lines">0</span>
                </div>
                <div class="stat">
                    <span>Level:</span>
                    <span class="stat-value" id="level">1</span>
                </div>
                <div class="stat">
                    <span>High Score:</span>
                    <span class="stat-value" id="highScore">0</span>
                </div>
            </div>

            <div class="panel-section">
                <h3>⏭️ Next Piece</h3>
                <canvas id="next-piece" width="100" height="100"></canvas>
            </div>

            <div class="panel-section">
                <h3>🏆 Top Scores</h3>
                <div class="high-scores">
                    <ol id="scoresList">
                        <li>Loading...</li>
                    </ol>
                </div>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script>
        let gameState = null;
        let gameLoop = null;
        let isPaused = false;
        let pieces = {};
        let colors = {};

        const canvas = document.getElementById('canvas');
        const ctx = canvas.getContext('2d');
        const nextCanvas = document.getElementById('next-piece');
        const nextCtx = nextCanvas.getContext('2d');

        const BLOCK_SIZE = 30;
        const NEXT_BLOCK_SIZE = 20;

        function initGame() {
            $.ajax({
                url: window.location.pathname,
                type: 'POST',
                data: { action: 'init' },
                success: function(response) {
                    const data = JSON.parse(response);
                    if (data.success) {
                        gameState = data.gameState;
                        pieces = data.pieces;
                        colors = data.colors;
                        $('#highScore').text(gameState.highScore);
                        renderGame();
                        startGameLoop();
                        loadHighScores();
                    }
                }
            });
        }

        function rotatePiece(piece, rotation) {
            const rotated = [];
            const n = piece.length;

            for (let r = 0; r < rotation % 4; r++) {
                const temp = [];
                for (let i = 0; i < n; i++) {
                    temp[i] = [];
                    for (let j = 0; j < n; j++) {
                        temp[i][j] = piece[n - 1 - j][i];
                    }
                }
                piece = temp;
            }
            return piece;
        }

        function renderGame() {
            if (!gameState) return;

            // Clear canvas
            ctx.fillStyle = '#000';
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            // Draw grid lines
            ctx.strokeStyle = '#222';
            for (let x = 0; x <= 10; x++) {
                ctx.beginPath();
                ctx.moveTo(x * BLOCK_SIZE, 0);
                ctx.lineTo(x * BLOCK_SIZE, 600);
                ctx.stroke();
            }
            for (let y = 0; y <= 20; y++) {
                ctx.beginPath();
                ctx.moveTo(0, y * BLOCK_SIZE);
                ctx.lineTo(300, y * BLOCK_SIZE);
                ctx.stroke();
            }

            // Draw locked blocks
            for (let y = 0; y < 20; y++) {
                for (let x = 0; x < 10; x++) {
                    if (gameState.board[y][x]) {
                        drawBlock(ctx, x, y, gameState.board[y][x], BLOCK_SIZE);
                    }
                }
            }

            // Draw current piece
            if (gameState.currentPiece) {
                const piece = pieces[gameState.currentPiece.type];
                const rotated = rotatePiece(piece, gameState.currentPiece.rotation);
                const color = colors[gameState.currentPiece.type];

                for (let y = 0; y < rotated.length; y++) {
                    for (let x = 0; x < rotated[y].length; x++) {
                        if (rotated[y][x]) {
                            drawBlock(ctx,
                                gameState.currentPiece.x + x,
                                gameState.currentPiece.y + y,
                                color,
                                BLOCK_SIZE
                            );
                        }
                    }
                }

                // Draw ghost piece (preview of where piece will land)
                drawGhostPiece();
            }

            // Draw next piece
            nextCtx.fillStyle = '#000';
            nextCtx.fillRect(0, 0, 100, 100);

            if (gameState.nextPiece) {
                const nextPiece = pieces[gameState.nextPiece];
                const nextColor = colors[gameState.nextPiece];
                const offsetX = (5 - nextPiece[0].length) / 2;
                const offsetY = (5 - nextPiece.length) / 2;

                for (let y = 0; y < nextPiece.length; y++) {
                    for (let x = 0; x < nextPiece[y].length; x++) {
                        if (nextPiece[y][x]) {
                            drawBlock(nextCtx, offsetX + x, offsetY + y, nextColor, NEXT_BLOCK_SIZE);
                        }
                    }
                }
            }

            // Update stats
            $('#score').text(gameState.score);
            $('#lines').text(gameState.lines);
            $('#level').text(gameState.level);

            // Show game over if needed
            if (gameState.gameOver) {
                $('#finalScore').text(gameState.score);
                $('#finalLines').text(gameState.lines);
                $('#gameOver').addClass('show');
                clearInterval(gameLoop);

                if (gameState.score > gameState.highScore) {
                    $('#highScore').text(gameState.score);
                }

                loadHighScores();
            }
        }

        function drawBlock(context, x, y, color, size) {
            const padding = 2;
            context.fillStyle = color;
            context.fillRect(
                x * size + padding,
                y * size + padding,
                size - padding * 2,
                size - padding * 2
            );

            // Add highlight for 3D effect
            context.fillStyle = 'rgba(255, 255, 255, 0.3)';
            context.fillRect(x * size + padding, y * size + padding, size - padding * 2, 3);
            context.fillRect(x * size + padding, y * size + padding, 3, size - padding * 2);

            // Add shadow for 3D effect
            context.fillStyle = 'rgba(0, 0, 0, 0.3)';
            context.fillRect(x * size + padding, y * size + size - padding - 3, size - padding * 2, 3);
            context.fillRect(x * size + size - padding - 3, y * size + padding, 3, size - padding * 2);
        }

        function drawGhostPiece() {
            if (!gameState || !gameState.currentPiece) return;

            const piece = pieces[gameState.currentPiece.type];
            const rotated = rotatePiece(piece, gameState.currentPiece.rotation);

            // Find lowest valid position
            let ghostY = gameState.currentPiece.y;
            while (isValidMove(rotated, gameState.currentPiece.x, ghostY + 1)) {
                ghostY++;
            }

            // Draw ghost
            ctx.fillStyle = 'rgba(255, 255, 255, 0.2)';
            for (let y = 0; y < rotated.length; y++) {
                for (let x = 0; x < rotated[y].length; x++) {
                    if (rotated[y][x]) {
                        ctx.fillRect(
                            (gameState.currentPiece.x + x) * BLOCK_SIZE + 2,
                            (ghostY + y) * BLOCK_SIZE + 2,
                            BLOCK_SIZE - 4,
                            BLOCK_SIZE - 4
                        );
                    }
                }
            }
        }

        function isValidMove(piece, x, y) {
            for (let py = 0; py < piece.length; py++) {
                for (let px = 0; px < piece[py].length; px++) {
                    if (piece[py][px]) {
                        const boardX = x + px;
                        const boardY = y + py;

                        if (boardX < 0 || boardX >= 10 || boardY >= 20) {
                            return false;
                        }

                        if (boardY >= 0 && gameState.board[boardY][boardX]) {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        function moveDown() {
            if (!gameState || gameState.gameOver || isPaused) return;

            const piece = gameState.currentPiece;
            const oldY = piece.y;

            // Try to move down
            piece.y += 1;

            // Check if move is valid
            const pieceShape = pieces[piece.type];
            if (!isValidMove(pieceShape, piece.x, piece.y, piece.rotation)) {
                // Revert move
                piece.y = oldY;

                // Lock piece to board
                lockPiece();

                // Clear completed lines
                const linesCleared = clearLines();
                if (linesCleared > 0) {
                    gameState.lines += linesCleared;
                    const scores = [0, 100, 300, 500, 800];
                    gameState.score += scores[linesCleared] * gameState.level;
                    gameState.level = 1 + Math.floor(gameState.lines / 10);
                }

                // Spawn new piece
                spawnNewPiece();

                // Check game over
                if (!isValidMove(pieces[gameState.currentPiece.type], gameState.currentPiece.x, gameState.currentPiece.y, gameState.currentPiece.rotation)) {
                    gameState.gameOver = true;
                    saveHighScore();
                }

                // Adjust speed for new level
                if (!gameState.gameOver) {
                    startGameLoop();
                }
            }

            renderGame();
        }

        function movePiece(direction) {
            if (!gameState || gameState.gameOver) return;

            const piece = gameState.currentPiece;
            const oldX = piece.x;
            const oldY = piece.y;
            const oldRotation = piece.rotation;

            // Apply movement
            if (direction === 'left') {
                piece.x -= 1;
            } else if (direction === 'right') {
                piece.x += 1;
            } else if (direction === 'down') {
                piece.y += 1;
            } else if (direction === 'rotate') {
                piece.rotation = (piece.rotation + 1) % 4;
            }

            // Check if move is valid
            const pieceShape = pieces[piece.type];
            if (!isValidMove(pieceShape, piece.x, piece.y, piece.rotation)) {
                // Revert move
                piece.x = oldX;
                piece.y = oldY;
                piece.rotation = oldRotation;
            }

            renderGame();
        }

        function hardDrop() {
            if (!gameState || gameState.gameOver) return;

            const piece = gameState.currentPiece;
            const pieceShape = pieces[piece.type];

            // Move down until invalid
            let dropDistance = 0;
            while (isValidMove(pieceShape, piece.x, piece.y + 1, piece.rotation)) {
                piece.y += 1;
                dropDistance += 1;
            }

            // Add bonus points
            gameState.score += dropDistance * 2;

            // Lock piece
            lockPiece();

            // Clear lines
            const linesCleared = clearLines();
            if (linesCleared > 0) {
                gameState.lines += linesCleared;
                const scores = [0, 100, 300, 500, 800];
                gameState.score += scores[linesCleared] * gameState.level;
                gameState.level = 1 + Math.floor(gameState.lines / 10);
            }

            // Spawn new piece
            spawnNewPiece();

            // Check game over
            if (!isValidMove(pieces[gameState.currentPiece.type], gameState.currentPiece.x, gameState.currentPiece.y, gameState.currentPiece.rotation)) {
                gameState.gameOver = true;
                saveHighScore();
            }

            // Adjust speed
            if (!gameState.gameOver) {
                startGameLoop();
            }

            renderGame();
        }

        function isValidMove(piece, x, y, rotation) {
            const rotated = rotatePiece(piece, rotation);

            for (let py = 0; py < rotated.length; py++) {
                for (let px = 0; px < rotated[py].length; px++) {
                    if (rotated[py][px]) {
                        const boardX = x + px;
                        const boardY = y + py;

                        if (boardX < 0 || boardX >= 10 || boardY >= 20) {
                            return false;
                        }

                        if (boardY >= 0 && gameState.board[boardY][boardX]) {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        function lockPiece() {
            const piece = gameState.currentPiece;
            const pieceShape = pieces[piece.type];
            const rotated = rotatePiece(pieceShape, piece.rotation);
            const color = colors[piece.type];

            for (let y = 0; y < rotated.length; y++) {
                for (let x = 0; x < rotated[y].length; x++) {
                    if (rotated[y][x]) {
                        const boardY = piece.y + y;
                        const boardX = piece.x + x;
                        if (boardY >= 0 && boardY < 20 && boardX >= 0 && boardX < 10) {
                            gameState.board[boardY][boardX] = color;
                        }
                    }
                }
            }
        }

        function clearLines() {
            let linesCleared = 0;
            let y = 19;

            while (y >= 0) {
                let lineFull = true;
                for (let x = 0; x < 10; x++) {
                    if (!gameState.board[y][x]) {
                        lineFull = false;
                        break;
                    }
                }

                if (lineFull) {
                    gameState.board.splice(y, 1);
                    gameState.board.unshift([0,0,0,0,0,0,0,0,0,0]);
                    linesCleared++;
                } else {
                    y--;
                }
            }

            return linesCleared;
        }

        function spawnNewPiece() {
            const pieceTypes = ['I', 'O', 'T', 'S', 'Z', 'J', 'L'];
            gameState.currentPiece = {
                type: gameState.nextPiece,
                x: 3,
                y: 0,
                rotation: 0
            };
            gameState.nextPiece = pieceTypes[Math.floor(Math.random() * pieceTypes.length)];
        }

        function saveHighScore() {
            if (gameState.score > gameState.highScore) {
                $.ajax({
                    url: window.location.pathname,
                    type: 'POST',
                    data: {
                        action: 'saveHighScore',
                        score: gameState.score
                    },
                    success: function(response) {
                        const data = JSON.parse(response);
                        if (data.success) {
                            gameState.highScore = gameState.score;
                            loadHighScores();
                        }
                    }
                });
            }
        }

        function startGameLoop() {
            clearInterval(gameLoop);
            const speed = Math.max(100, 800 - (gameState.level - 1) * 50);
            gameLoop = setInterval(moveDown, speed);
        }

        function startNewGame() {
            $('#gameOver').removeClass('show');
            isPaused = false;
            initGame();
        }

        function togglePause() {
            if (gameState && !gameState.gameOver) {
                isPaused = !isPaused;
                if (isPaused) {
                    clearInterval(gameLoop);
                } else {
                    startGameLoop();
                }
            }
        }

        function loadHighScores() {
            $.ajax({
                url: window.location.pathname,
                type: 'POST',
                data: { action: 'getHighScores' },
                success: function(response) {
                    const data = JSON.parse(response);
                    if (data.success && data.scores && data.scores.length > 0) {
                        let html = '';
                        data.scores.forEach(function(score, index) {
                            const name = score.name || 'Unknown';
                            const scoreValue = score.score || 0;

                            // Highlight current user's scores
                            const isCurrentUser = data.playerName && name === data.playerName;
                            const style = isCurrentUser ? ' style="font-weight: bold; color: #fff;"' : '';

                            html += '<li' + style + '>';
                            html += name.split(' ')[0] + ': ' + scoreValue;
                            html += '</li>';
                        });
                        $('#scoresList').html(html);
                    } else {
                        $('#scoresList').html('<li>No scores yet - be the first!</li>');
                    }
                }
            });
        }

        // Keyboard controls
        $(document).keydown(function(e) {
            if (!gameState || gameState.gameOver) return;

            switch(e.keyCode) {
                case 37: // Left
                    movePiece('left');
                    e.preventDefault();
                    break;
                case 39: // Right
                    movePiece('right');
                    e.preventDefault();
                    break;
                case 40: // Down (soft drop)
                    movePiece('down');
                    e.preventDefault();
                    break;
                case 38: // Up (rotate)
                    movePiece('rotate');
                    e.preventDefault();
                    break;
                case 32: // Space (hard drop)
                    hardDrop();
                    e.preventDefault();
                    break;
                case 80: // P (pause)
                    togglePause();
                    e.preventDefault();
                    break;
            }
        });

        // Start game on load
        $(document).ready(function() {
            initGame();
        });
    </script>
</body>
</html>
'''
