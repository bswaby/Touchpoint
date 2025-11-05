"""
TouchPoint Breakout Game
Classic brick-breaking game using IronPython 2.7.3 and HTML5 Canvas
Access via: /PyScriptForm/BreakoutGame
"""

import json
import datetime

# Handle high score submission
if model.HttpMethod == "post" and hasattr(Data, 'action'):
    action = Data.action

    if action == 'submit_score':
        # Load existing high scores
        try:
            scores_json = model.TextContent("BreakoutHighScores")
            scores = json.loads(scores_json) if scores_json else []
        except:
            scores = []

        # Add new score
        new_score = {
            'player': Data.player if hasattr(Data, 'player') else 'Anonymous',
            'score': int(Data.score),
            'level': int(Data.level),
            'bricks': int(Data.bricks),
            'date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        }

        scores.append(new_score)

        # Sort by score (descending) and keep top 10
        scores = sorted(scores, key=lambda x: x['score'], reverse=True)[:10]

        # Save back to content storage
        model.WriteContentText("BreakoutHighScores", json.dumps(scores), "")

        # Return updated leaderboard
        print json.dumps({'success': True, 'scores': scores})
        return

    elif action == 'get_scores':
        # Return current high scores
        try:
            scores_json = model.TextContent("BreakoutHighScores")
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
    <title>TouchPoint Breakout</title>
    <style>
        body {
            margin: 0;
            padding: 20px;
            font-family: 'Courier New', monospace;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 50%, #7e22ce 100%);
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
            box-shadow: 0 0 30px rgba(255, 255, 255, 0.4);
            background: #000;
            cursor: none;
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

        .lives {
            display: flex;
            gap: 5px;
            margin-top: 10px;
        }

        .life {
            width: 20px;
            height: 20px;
            background: #ff4444;
            border-radius: 50%;
            box-shadow: 0 0 5px rgba(255, 68, 68, 0.5);
        }

        button {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
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

        .game-over-modal, .level-complete-modal {
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

        .game-over-modal.show, .level-complete-modal.show {
            display: block;
        }

        .game-over-modal h1 {
            color: #ff4444;
            margin: 0 0 20px 0;
            font-size: 48px;
        }

        .level-complete-modal h1 {
            color: #44ff44;
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

        .powerup-info {
            background: rgba(255, 255, 255, 0.05);
            padding: 10px;
            border-radius: 5px;
            margin-top: 10px;
            font-size: 12px;
        }

        .powerup-active {
            color: #44ff44;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <div class="game-container">
        <div class="game-area">
            <h1 style="margin: 0 0 20px 0; text-align: center;">🎯 BREAKOUT 🎯</h1>
            <div style="position: relative;">
                <canvas id="gameCanvas" width="600" height="600"></canvas>
                <div id="pauseOverlay">PAUSED</div>
            </div>
            <div class="controls">
                <button id="startBtn" onclick="startGame()">START GAME</button>
                <button id="pauseBtn" onclick="togglePause()" style="display:none;">PAUSE</button>
                <h3 style="margin-top: 20px;">Controls:</h3>
                <div>
                    <span class="key">←</span> Move Left
                    <span class="key">→</span> Move Right
                    <span class="key">Mouse</span> Move Paddle
                </div>
                <div>
                    <span class="key">Space</span> Launch Ball
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
                    <span class="stat-label">Bricks:</span>
                    <span class="stat-value" id="bricks">0</span>
                </div>
                <div style="margin-top: 15px;">
                    <span class="stat-label">Lives:</span>
                    <div class="lives" id="livesDisplay"></div>
                </div>
            </div>

            <div class="info-panel">
                <h2>Power-ups</h2>
                <div class="powerup-info">
                    <div>🟢 <strong>Expand:</strong> Wider paddle</div>
                    <div>🔵 <strong>Slow:</strong> Slower ball</div>
                    <div>🟡 <strong>Multi:</strong> Extra balls</div>
                    <div>🔴 <strong>Fire:</strong> Pierce bricks</div>
                </div>
                <div id="activePowerups" style="margin-top: 10px;"></div>
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
            <span class="stat-label">Bricks Destroyed:</span>
            <span class="stat-value" id="finalBricks">0</span>
        </div>
        <div style="margin-top: 30px;">
            <input type="text" id="playerName" placeholder="Enter your name" maxlength="20" />
            <br/>
            <button onclick="submitScore()">SUBMIT SCORE</button>
            <button onclick="restartGame()">PLAY AGAIN</button>
        </div>
    </div>

    <div id="levelCompleteModal" class="level-complete-modal">
        <h1>LEVEL COMPLETE!</h1>
        <div class="stat">
            <span class="stat-label">Level:</span>
            <span class="stat-value" id="completedLevel">1</span>
        </div>
        <div class="stat">
            <span class="stat-label">Bonus:</span>
            <span class="stat-value" id="levelBonus">0</span>
        </div>
        <div style="margin-top: 30px;">
            <button onclick="nextLevel()">NEXT LEVEL</button>
        </div>
    </div>

    <script>
        // Game constants
        const canvas = document.getElementById('gameCanvas');
        const ctx = canvas.getContext('2d');

        const PADDLE_WIDTH = 100;
        const PADDLE_HEIGHT = 15;
        const BALL_RADIUS = 8;
        const BRICK_ROWS = 5;
        const BRICK_COLS = 10;
        const BRICK_WIDTH = 56;
        const BRICK_HEIGHT = 20;
        const BRICK_PADDING = 4;
        const BRICK_OFFSET_TOP = 60;
        const BRICK_OFFSET_LEFT = 10;

        // Brick colors by row
        const BRICK_COLORS = [
            '#FF0D72',
            '#FF8E0D',
            '#FFE138',
            '#0DFF72',
            '#0DC2FF'
        ];

        // Brick values by row
        const BRICK_VALUES = [50, 40, 30, 20, 10];

        // Power-up types
        const POWERUP_TYPES = [
            { type: 'expand', color: '#44ff44', symbol: '🟢' },
            { type: 'slow', color: '#4444ff', symbol: '🔵' },
            { type: 'multi', color: '#ffff44', symbol: '🟡' },
            { type: 'fire', color: '#ff4444', symbol: '🔴' }
        ];

        // Game state
        let paddle = { x: canvas.width / 2 - PADDLE_WIDTH / 2, y: canvas.height - 40, width: PADDLE_WIDTH, height: PADDLE_HEIGHT };
        let balls = [];
        let bricks = [];
        let powerups = [];
        let activePowerups = [];
        let score = 0;
        let level = 1;
        let totalBricksDestroyed = 0;
        let lives = 3;
        let gameRunning = false;
        let gamePaused = false;
        let ballLaunched = false;

        // Mouse control
        let mouseX = canvas.width / 2;

        canvas.addEventListener('mousemove', (e) => {
            const rect = canvas.getBoundingClientRect();
            mouseX = e.clientX - rect.left;
        });

        // Initialize game
        function createBricks() {
            bricks = [];
            for (let row = 0; row < BRICK_ROWS; row++) {
                for (let col = 0; col < BRICK_COLS; col++) {
                    bricks.push({
                        x: col * (BRICK_WIDTH + BRICK_PADDING) + BRICK_OFFSET_LEFT,
                        y: row * (BRICK_HEIGHT + BRICK_PADDING) + BRICK_OFFSET_TOP,
                        status: 1,
                        color: BRICK_COLORS[row],
                        value: BRICK_VALUES[row],
                        hits: 0,
                        maxHits: Math.min(level, 3)
                    });
                }
            }
        }

        function createBall(x, y, dx, dy) {
            return {
                x: x || paddle.x + paddle.width / 2,
                y: y || paddle.y - BALL_RADIUS,
                dx: dx || 0,
                dy: dy || 0,
                radius: BALL_RADIUS,
                speed: 3 + level * 0.5
            };
        }

        function resetPaddle() {
            paddle.x = canvas.width / 2 - paddle.width / 2;
            paddle.width = PADDLE_WIDTH;
        }

        function startGame() {
            score = 0;
            level = 1;
            lives = 3;
            totalBricksDestroyed = 0;
            activePowerups = [];
            powerups = [];

            resetPaddle();
            createBricks();
            balls = [createBall()];
            ballLaunched = false;

            gameRunning = true;
            gamePaused = false;

            document.getElementById('startBtn').style.display = 'none';
            document.getElementById('pauseBtn').style.display = 'inline-block';
            document.getElementById('gameOverModal').classList.remove('show');

            updateStats();
            gameLoop();
        }

        function nextLevel() {
            level++;
            resetPaddle();
            createBricks();
            balls = [createBall()];
            ballLaunched = false;
            activePowerups = [];
            powerups = [];

            document.getElementById('levelCompleteModal').classList.remove('show');
            gamePaused = false;
        }

        function updateStats() {
            document.getElementById('score').textContent = score;
            document.getElementById('level').textContent = level;
            document.getElementById('bricks').textContent = totalBricksDestroyed;

            // Update lives display
            const livesDisplay = document.getElementById('livesDisplay');
            livesDisplay.innerHTML = '';
            for (let i = 0; i < lives; i++) {
                const life = document.createElement('div');
                life.className = 'life';
                livesDisplay.appendChild(life);
            }

            // Update active powerups display
            const powerupsDisplay = document.getElementById('activePowerups');
            powerupsDisplay.innerHTML = '';
            activePowerups.forEach(p => {
                const div = document.createElement('div');
                div.className = 'powerup-active';
                div.textContent = p.symbol + ' ' + p.type.toUpperCase() + ' active';
                powerupsDisplay.appendChild(div);
            });
        }

        // Drawing functions
        function drawPaddle() {
            ctx.fillStyle = '#fff';
            ctx.fillRect(paddle.x, paddle.y, paddle.width, paddle.height);

            // Glow effect
            ctx.shadowBlur = 15;
            ctx.shadowColor = '#fff';
            ctx.fillRect(paddle.x, paddle.y, paddle.width, paddle.height);
            ctx.shadowBlur = 0;
        }

        function drawBall(ball) {
            // Check if fire powerup is active
            const hasFire = activePowerups.some(p => p.type === 'fire');

            ctx.beginPath();
            ctx.arc(ball.x, ball.y, ball.radius, 0, Math.PI * 2);
            ctx.fillStyle = hasFire ? '#ff4444' : '#fff';
            ctx.fill();
            ctx.closePath();

            // Glow effect
            ctx.shadowBlur = hasFire ? 20 : 10;
            ctx.shadowColor = hasFire ? '#ff4444' : '#fff';
            ctx.beginPath();
            ctx.arc(ball.x, ball.y, ball.radius, 0, Math.PI * 2);
            ctx.fill();
            ctx.closePath();
            ctx.shadowBlur = 0;
        }

        function drawBricks() {
            bricks.forEach(brick => {
                if (brick.status === 1) {
                    // Adjust opacity based on hits
                    const opacity = 1 - (brick.hits / brick.maxHits) * 0.5;
                    ctx.globalAlpha = opacity;

                    ctx.fillStyle = brick.color;
                    ctx.fillRect(brick.x, brick.y, BRICK_WIDTH, BRICK_HEIGHT);

                    // Border
                    ctx.strokeStyle = '#000';
                    ctx.lineWidth = 2;
                    ctx.strokeRect(brick.x, brick.y, BRICK_WIDTH, BRICK_HEIGHT);

                    ctx.globalAlpha = 1;
                }
            });
        }

        function drawPowerups() {
            powerups.forEach(powerup => {
                ctx.font = '20px Arial';
                ctx.fillText(powerup.symbol, powerup.x, powerup.y);
            });
        }

        function draw() {
            // Clear canvas
            ctx.fillStyle = '#000';
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            drawBricks();
            drawPaddle();
            balls.forEach(ball => drawBall(ball));
            drawPowerups();

            // Draw score at top
            ctx.fillStyle = '#fff';
            ctx.font = '20px Courier New';
            ctx.fillText('Score: ' + score, 10, 30);
            ctx.fillText('Level: ' + level, canvas.width - 100, 30);
        }

        // Game logic
        function updatePaddle() {
            paddle.x = mouseX - paddle.width / 2;

            // Keep paddle on screen
            if (paddle.x < 0) paddle.x = 0;
            if (paddle.x + paddle.width > canvas.width) {
                paddle.x = canvas.width - paddle.width;
            }
        }

        function updateBalls() {
            balls.forEach((ball, index) => {
                if (!ballLaunched) {
                    // Ball follows paddle
                    ball.x = paddle.x + paddle.width / 2;
                    ball.y = paddle.y - BALL_RADIUS;
                } else {
                    // Move ball
                    ball.x += ball.dx;
                    ball.y += ball.dy;

                    // Wall collision
                    if (ball.x + ball.radius > canvas.width || ball.x - ball.radius < 0) {
                        ball.dx = -ball.dx;
                    }

                    if (ball.y - ball.radius < 0) {
                        ball.dy = -ball.dy;
                    }

                    // Paddle collision
                    if (ball.y + ball.radius > paddle.y &&
                        ball.y + ball.radius < paddle.y + paddle.height &&
                        ball.x > paddle.x &&
                        ball.x < paddle.x + paddle.width) {

                        // Angle based on where it hits the paddle
                        const hitPos = (ball.x - paddle.x) / paddle.width;
                        const angle = (hitPos - 0.5) * Math.PI * 0.6;

                        const speed = Math.sqrt(ball.dx * ball.dx + ball.dy * ball.dy);
                        ball.dx = speed * Math.sin(angle);
                        ball.dy = -Math.abs(speed * Math.cos(angle));
                    }

                    // Bottom - lose life
                    if (ball.y + ball.radius > canvas.height) {
                        balls.splice(index, 1);

                        if (balls.length === 0) {
                            lives--;
                            if (lives <= 0) {
                                gameOver();
                            } else {
                                ballLaunched = false;
                                balls = [createBall()];
                            }
                            updateStats();
                        }
                    }
                }
            });
        }

        function updateBricks() {
            const hasFire = activePowerups.some(p => p.type === 'fire');

            balls.forEach(ball => {
                bricks.forEach(brick => {
                    if (brick.status === 1) {
                        if (ball.x + ball.radius > brick.x &&
                            ball.x - ball.radius < brick.x + BRICK_WIDTH &&
                            ball.y + ball.radius > brick.y &&
                            ball.y - ball.radius < brick.y + BRICK_HEIGHT) {

                            if (!hasFire) {
                                ball.dy = -ball.dy;
                            }

                            brick.hits++;

                            if (brick.hits >= brick.maxHits) {
                                brick.status = 0;
                                score += brick.value * level;
                                totalBricksDestroyed++;

                                // Random powerup drop (10% chance)
                                if (Math.random() < 0.1) {
                                    const powerupType = POWERUP_TYPES[Math.floor(Math.random() * POWERUP_TYPES.length)];
                                    powerups.push({
                                        x: brick.x + BRICK_WIDTH / 2,
                                        y: brick.y + BRICK_HEIGHT / 2,
                                        dy: 2,
                                        type: powerupType.type,
                                        color: powerupType.color,
                                        symbol: powerupType.symbol
                                    });
                                }

                                // Check level complete
                                if (bricks.filter(b => b.status === 1).length === 0) {
                                    levelComplete();
                                }
                            }

                            updateStats();
                        }
                    }
                });
            });
        }

        function updatePowerups() {
            powerups.forEach((powerup, index) => {
                powerup.y += powerup.dy;

                // Check paddle collision
                if (powerup.y > paddle.y &&
                    powerup.y < paddle.y + paddle.height &&
                    powerup.x > paddle.x &&
                    powerup.x < paddle.x + paddle.width) {

                    activatePowerup(powerup);
                    powerups.splice(index, 1);
                }

                // Remove if off screen
                if (powerup.y > canvas.height) {
                    powerups.splice(index, 1);
                }
            });
        }

        function activatePowerup(powerup) {
            switch (powerup.type) {
                case 'expand':
                    paddle.width = PADDLE_WIDTH * 1.5;
                    addTimedPowerup(powerup, 10000, () => {
                        paddle.width = PADDLE_WIDTH;
                    });
                    break;

                case 'slow':
                    balls.forEach(ball => {
                        ball.speed *= 0.7;
                        const currentSpeed = Math.sqrt(ball.dx * ball.dx + ball.dy * ball.dy);
                        const factor = ball.speed / currentSpeed;
                        ball.dx *= factor;
                        ball.dy *= factor;
                    });
                    addTimedPowerup(powerup, 8000, () => {
                        balls.forEach(ball => {
                            ball.speed /= 0.7;
                            const currentSpeed = Math.sqrt(ball.dx * ball.dx + ball.dy * ball.dy);
                            const factor = ball.speed / currentSpeed;
                            ball.dx *= factor;
                            ball.dy *= factor;
                        });
                    });
                    break;

                case 'multi':
                    if (ballLaunched && balls.length > 0) {
                        const mainBall = balls[0];
                        balls.push(createBall(mainBall.x, mainBall.y, -mainBall.dx, mainBall.dy));
                        balls.push(createBall(mainBall.x, mainBall.y, mainBall.dx * 0.7, -mainBall.dy * 0.7));
                    }
                    break;

                case 'fire':
                    addTimedPowerup(powerup, 10000);
                    break;
            }

            updateStats();
        }

        function addTimedPowerup(powerup, duration, onExpire) {
            activePowerups.push(powerup);

            setTimeout(() => {
                const index = activePowerups.indexOf(powerup);
                if (index > -1) {
                    activePowerups.splice(index, 1);
                    if (onExpire) onExpire();
                    updateStats();
                }
            }, duration);
        }

        function levelComplete() {
            gamePaused = true;

            const bonus = lives * 1000 + bricks.filter(b => b.status === 1).length * 50;
            score += bonus;

            document.getElementById('completedLevel').textContent = level;
            document.getElementById('levelBonus').textContent = bonus;
            document.getElementById('levelCompleteModal').classList.add('show');

            updateStats();
        }

        function gameOver() {
            gameRunning = false;

            document.getElementById('finalScore').textContent = score;
            document.getElementById('finalLevel').textContent = level;
            document.getElementById('finalBricks').textContent = totalBricksDestroyed;
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

        function togglePause() {
            gamePaused = !gamePaused;
            document.getElementById('pauseOverlay').classList.toggle('show');
            document.getElementById('pauseBtn').textContent = gamePaused ? 'RESUME' : 'PAUSE';
        }

        function submitScore() {
            let playerName = document.getElementById('playerName').value.trim() || 'Anonymous';

            var pathname = window.location.pathname;
            var scriptName = pathname.split('/').pop().split('?')[0];

            fetch('/PyScriptForm/' + scriptName, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: 'action=submit_score&player=' + encodeURIComponent(playerName) +
                      '&score=' + score + '&level=' + level + '&bricks=' + totalBricksDestroyed
            })
            .then(response => response.json())
            .then(data => {
                loadLeaderboard();
                document.getElementById('gameOverModal').classList.remove('show');
            })
            .catch(error => console.error('Error:', error));
        }

        function loadLeaderboard() {
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
            if (e.key === ' ') {
                e.preventDefault();
                if (gameRunning && !gamePaused && !ballLaunched) {
                    ballLaunched = true;
                    balls[0].dx = 3 + level * 0.3;
                    balls[0].dy = -(3 + level * 0.3);
                }
            }

            if (e.key === 'p' || e.key === 'P') {
                if (gameRunning) togglePause();
            }

            if (e.key === 'ArrowLeft') {
                mouseX = Math.max(paddle.width / 2, mouseX - 30);
            }

            if (e.key === 'ArrowRight') {
                mouseX = Math.min(canvas.width - paddle.width / 2, mouseX + 30);
            }
        });

        document.getElementById('playerName').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                submitScore();
            }
        });

        // Game loop
        function gameLoop() {
            if (!gameRunning) return;

            if (!gamePaused) {
                updatePaddle();
                updateBalls();
                updateBricks();
                updatePowerups();
            }

            draw();
            requestAnimationFrame(gameLoop);
        }

        // Load leaderboard on page load
        loadLeaderboard();

        // Initial draw
        draw();
    </script>
</body>
</html>
'''
