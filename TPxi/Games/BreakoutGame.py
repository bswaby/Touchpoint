#Roles=Access
# Breakout Game
# Create By: Ben Swaby
# Email: bswaby@fbchtn.org

import json
import random
import datetime
import math

model.Header = 'Breakout Game'

# Power-up types
POWERUPS = {
    'expand': {'color': '#00ff00', 'duration': 10000, 'effect': 'Wider Paddle'},
    'slow': {'color': '#00ffff', 'duration': 10000, 'effect': 'Slower Ball'},
    'multi': {'color': '#ff00ff', 'duration': 0, 'effect': 'Multi-Ball'},
    'fire': {'color': '#ff0000', 'duration': 10000, 'effect': 'Fire Ball'}
}

# Handle AJAX requests
if model.HttpMethod == "post" and hasattr(Data, 'action'):
    action = Data.action

    if action == 'init':
        # Initialize new game
        level = int(Data.level) if hasattr(Data, 'level') else 1

        # Create bricks
        bricks = []
        rows = min(5 + level, 10)
        colors = ['#ff0000', '#ff7700', '#ffff00', '#00ff00', '#0000ff', '#4b0082', '#9400d3']

        for row in range(rows):
            for col in range(10):
                # Some bricks may have powerups (10% chance)
                has_powerup = random.random() < 0.1
                powerup_type = random.choice(['expand', 'slow', 'multi', 'fire']) if has_powerup else None

                bricks.append({
                    'x': col * 60 + 5,
                    'y': row * 20 + 50,
                    'width': 55,
                    'height': 15,
                    'color': colors[row % len(colors)],
                    'hits': 1 if row < 3 else (2 if row < 6 else 3),
                    'maxHits': 1 if row < 3 else (2 if row < 6 else 3),
                    'powerup': powerup_type
                })

        game_state = {
            'paddle': {
                'x': 250,
                'y': 550,
                'width': 100,
                'height': 15,
                'speed': 8
            },
            'balls': [{
                'x': 300,
                'y': 535,
                'dx': 3,
                'dy': -3,
                'radius': 8,
                'speed': 3
            }],
            'bricks': bricks,
            'powerups': [],
            'activePowerups': {},
            'lives': 3,
            'score': 0,
            'level': level,
            'gameOver': False,
            'levelComplete': False,
            'highScore': 0
        }

        # Get user's best score from leaderboard
        try:
            person = model.GetPerson(model.UserPeopleId)
            player_name = person.Name if person else 'Unknown Player'

            scores_json = model.TextContent('BreakoutHighScores')
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

        print json.dumps({'success': True, 'gameState': game_state})

    elif action == 'saveHighScore':
        # Save high score to shared leaderboard
        score = int(Data.score) if hasattr(Data, 'score') else 0

        try:
            # Get current user's name
            person = model.GetPerson(model.UserPeopleId)
            player_name = person.Name if person else 'Unknown Player'

            # Load existing scores
            try:
                scores_json = model.TextContent('BreakoutHighScores')
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
            model.WriteContentText('BreakoutHighScores', json.dumps(scores), '')

            print json.dumps({'success': True, 'playerName': player_name})
        except Exception as e:
            print json.dumps({'success': False, 'error': str(e)})

    elif action == 'getHighScores':
        # Get top 10 high scores from content storage
        try:
            scores_json = model.TextContent('BreakoutHighScores')
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
    <title>Breakout Game</title>
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
            background: linear-gradient(135deg, #232526 0%, #414345 100%);
            padding: 20px;
        }

        .game-container {
            display: flex;
            gap: 20px;
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            padding: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.5);
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
            color: #232526;
            margin-bottom: 20px;
            font-size: 36px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }

        #game-board {
            width: 600px;
            height: 600px;
            border: 3px solid #333;
            position: relative;
            background: linear-gradient(180deg, #000000 0%, #1a1a2e 100%);
            box-shadow: inset 0 0 30px rgba(0,0,0,0.8);
            border-radius: 10px;
            cursor: none;
        }

        canvas {
            display: block;
            width: 100%;
            height: 100%;
            border-radius: 7px;
        }

        .side-panel {
            display: flex;
            flex-direction: column;
            gap: 20px;
            min-width: 200px;
        }

        .panel-section {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
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

        .lives {
            display: flex;
            gap: 5px;
            justify-content: center;
            margin-top: 10px;
        }

        .life {
            width: 20px;
            height: 20px;
            background: #ff0000;
            border-radius: 50%;
            box-shadow: 0 2px 5px rgba(0,0,0,0.3);
        }

        .controls {
            margin-top: 20px;
            text-align: center;
        }

        button {
            padding: 12px 24px;
            font-size: 16px;
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
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

        .game-over, .level-complete {
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

        .game-over.show, .level-complete.show {
            display: block;
            animation: fadeIn 0.3s;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translate(-50%, -50%) scale(0.8); }
            to { opacity: 1; transform: translate(-50%, -50%) scale(1); }
        }

        .game-over h2, .level-complete h2 {
            margin-bottom: 20px;
            font-size: 32px;
        }

        .instructions {
            margin-top: 20px;
            color: #666;
            font-size: 14px;
            text-align: center;
        }

        .powerup-list {
            margin-top: 10px;
            font-size: 12px;
        }

        .powerup-item {
            margin: 5px 0;
            padding: 5px;
            background: rgba(255,255,255,0.1);
            border-radius: 5px;
        }

        .high-scores {
            max-height: 250px;
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
                padding: 15px;
                max-width: 95vw;
            }

            #game-board {
                width: min(500px, 90vw);
                height: min(500px, 90vw);
            }

            canvas {
                width: 100% !important;
                height: 100% !important;
            }

            .side-panel {
                flex-direction: row;
                flex-wrap: wrap;
                width: 100%;
            }

            .panel-section {
                flex: 1;
                min-width: 150px;
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
            <h1>🎯 Breakout</h1>

            <div id="game-board">
                <canvas id="canvas" width="600" height="600"></canvas>
                <div class="game-over" id="gameOver">
                    <h2>Game Over!</h2>
                    <p style="font-size: 20px; margin: 10px 0;">Final Score: <span id="finalScore">0</span></p>
                    <p style="font-size: 16px; margin: 10px 0;">Level Reached: <span id="finalLevel">0</span></p>
                    <button onclick="startNewGame()" style="margin-top: 20px;">Play Again</button>
                </div>
                <div class="level-complete" id="levelComplete">
                    <h2>Level Complete!</h2>
                    <p style="font-size: 20px; margin: 10px 0;">Score: <span id="levelScore">0</span></p>
                    <button onclick="nextLevel()" style="margin-top: 20px;">Next Level</button>
                </div>
            </div>

            <div class="controls">
                <button onclick="startNewGame()">New Game</button>
                <button onclick="togglePause()">Pause</button>
            </div>

            <div class="instructions">
                <p>Move mouse to control paddle | Break all bricks!</p>
                <p style="margin-top: 10px; font-size: 12px;">
                    <span style="color: #00ff00;">🟢 Expand</span> |
                    <span style="color: #00ffff;">🔵 Slow</span> |
                    <span style="color: #ff00ff;">🟣 Multi-Ball</span> |
                    <span style="color: #ff0000;">🔴 Fire</span>
                </p>
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
                    <span>Level:</span>
                    <span class="stat-value" id="level">1</span>
                </div>
                <div class="stat">
                    <span>High Score:</span>
                    <span class="stat-value" id="highScore">0</span>
                </div>
                <div>
                    <p style="margin-top: 10px;">Lives:</p>
                    <div class="lives" id="lives"></div>
                </div>
            </div>

            <div class="panel-section">
                <h3>⚡ Active Power-ups</h3>
                <div class="powerup-list" id="powerupList">
                    <p style="font-size: 12px; opacity: 0.7;">Collect falling power-ups!</p>
                </div>
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
        let mouseX = 300;
        let lastUpdateTime = 0;
        let fps = 0;

        const canvas = document.getElementById('canvas');
        const ctx = canvas.getContext('2d');

        // Ensure canvas is properly sized
        canvas.width = 600;
        canvas.height = 600;

        function initGame(level) {
            level = level || 1;
            $.ajax({
                url: window.location.pathname,
                type: 'POST',
                data: { action: 'init', level: level },
                success: function(response) {
                    const data = JSON.parse(response);
                    if (data.success) {
                        gameState = data.gameState;
                        $('#highScore').text(gameState.highScore);
                        renderGame();
                        startGameLoop();
                        loadHighScores();
                    }
                }
            });
        }

        function renderGame() {
            if (!gameState) return;

            // Clear canvas with gradient
            const gradient = ctx.createLinearGradient(0, 0, 0, 600);
            gradient.addColorStop(0, '#000000');
            gradient.addColorStop(1, '#1a1a2e');
            ctx.fillStyle = gradient;
            ctx.fillRect(0, 0, 600, 600);

            // Draw bricks
            gameState.bricks.forEach(function(brick) {
                // Brick color based on hits
                const alpha = brick.hits / brick.maxHits;
                ctx.fillStyle = brick.color;
                ctx.globalAlpha = 0.3 + (alpha * 0.7);

                ctx.fillRect(brick.x, brick.y, brick.width, brick.height);

                // Brick border
                ctx.strokeStyle = '#fff';
                ctx.lineWidth = 1;
                ctx.globalAlpha = 1;
                ctx.strokeRect(brick.x, brick.y, brick.width, brick.height);

                // Powerup indicator
                if (brick.powerup) {
                    ctx.fillStyle = '#ffffff';
                    ctx.font = '10px Arial';
                    ctx.fillText('⚡', brick.x + brick.width/2 - 5, brick.y + brick.height/2 + 3);
                }
            });

            ctx.globalAlpha = 1;

            // Draw paddle
            const paddle = gameState.paddle;
            const paddleGradient = ctx.createLinearGradient(paddle.x, paddle.y, paddle.x, paddle.y + paddle.height);
            paddleGradient.addColorStop(0, '#00ffff');
            paddleGradient.addColorStop(1, '#0088ff');
            ctx.fillStyle = paddleGradient;
            ctx.fillRect(paddle.x, paddle.y, paddle.width, paddle.height);

            // Paddle glow
            ctx.shadowBlur = 15;
            ctx.shadowColor = '#00ffff';
            ctx.fillRect(paddle.x, paddle.y, paddle.width, paddle.height);
            ctx.shadowBlur = 0;

            // Draw balls
            gameState.balls.forEach(function(ball) {
                const isFire = gameState.activePowerups.fire;

                ctx.beginPath();
                ctx.arc(ball.x, ball.y, ball.radius, 0, Math.PI * 2);

                if (isFire) {
                    const ballGradient = ctx.createRadialGradient(ball.x, ball.y, 0, ball.x, ball.y, ball.radius);
                    ballGradient.addColorStop(0, '#ffff00');
                    ballGradient.addColorStop(0.5, '#ff8800');
                    ballGradient.addColorStop(1, '#ff0000');
                    ctx.fillStyle = ballGradient;
                } else {
                    const ballGradient = ctx.createRadialGradient(ball.x, ball.y, 0, ball.x, ball.y, ball.radius);
                    ballGradient.addColorStop(0, '#ffffff');
                    ballGradient.addColorStop(1, '#aaaaaa');
                    ctx.fillStyle = ballGradient;
                }

                ctx.fill();

                // Ball glow
                ctx.shadowBlur = 10;
                ctx.shadowColor = isFire ? '#ff0000' : '#ffffff';
                ctx.fill();
                ctx.shadowBlur = 0;
            });

            // Draw powerups
            gameState.powerups.forEach(function(powerup) {
                ctx.fillStyle = getPowerupColor(powerup.type);
                ctx.fillRect(powerup.x, powerup.y, powerup.width, powerup.height);

                ctx.fillStyle = '#ffffff';
                ctx.font = '16px Arial';
                ctx.fillText(getPowerupIcon(powerup.type), powerup.x + 7, powerup.y + 22);
            });

            // Update UI
            $('#score').text(gameState.score);
            $('#level').text(gameState.level);

            // Update lives display
            let livesHtml = '';
            for (let i = 0; i < gameState.lives; i++) {
                livesHtml += '<div class="life"></div>';
            }
            $('#lives').html(livesHtml);

            // Update active powerups display
            let powerupsHtml = '';
            for (let ptype in gameState.activePowerups) {
                powerupsHtml += '<div class="powerup-item">' + getPowerupName(ptype) + '</div>';
            }
            if (powerupsHtml === '') {
                powerupsHtml = '<p style="font-size: 12px; opacity: 0.7;">None active</p>';
            }
            $('#powerupList').html(powerupsHtml);

            // Check game over
            if (gameState.gameOver) {
                $('#finalScore').text(gameState.score);
                $('#finalLevel').text(gameState.level);
                $('#gameOver').addClass('show');
                clearInterval(gameLoop);

                if (gameState.score > gameState.highScore) {
                    $('#highScore').text(gameState.score);
                }

                loadHighScores();
            }

            // Check level complete
            if (gameState.levelComplete) {
                $('#levelScore').text(gameState.score);
                $('#levelComplete').addClass('show');
                clearInterval(gameLoop);
            }
        }

        function getPowerupColor(type) {
            const colors = {
                'expand': '#00ff00',
                'slow': '#00ffff',
                'multi': '#ff00ff',
                'fire': '#ff0000'
            };
            return colors[type] || '#ffffff';
        }

        function getPowerupIcon(type) {
            const icons = {
                'expand': '↔',
                'slow': '🐌',
                'multi': '●●',
                'fire': '🔥'
            };
            return icons[type] || '?';
        }

        function getPowerupName(type) {
            const names = {
                'expand': 'Wide Paddle',
                'slow': 'Slow Ball',
                'multi': 'Multi-Ball',
                'fire': 'Fire Ball'
            };
            return names[type] || 'Unknown';
        }

        function updateGame() {
            if (!gameState || gameState.gameOver || gameState.levelComplete || isPaused) return;

            // Update paddle position based on mouse
            const paddleWidth = gameState.paddle.width;
            gameState.paddle.x = Math.max(0, Math.min(600 - paddleWidth, mouseX - paddleWidth / 2));

            // Update balls
            for (let i = gameState.balls.length - 1; i >= 0; i--) {
                const ball = gameState.balls[i];
                ball.x += ball.dx;
                ball.y += ball.dy;

                // Wall collision
                if (ball.x - ball.radius <= 0 || ball.x + ball.radius >= 600) {
                    ball.dx = -ball.dx;
                }

                // Ceiling collision
                if (ball.y - ball.radius <= 0) {
                    ball.dy = -ball.dy;
                }

                // Paddle collision
                const paddle = gameState.paddle;
                if (ball.y + ball.radius >= paddle.y &&
                    ball.y - ball.radius <= paddle.y + paddle.height &&
                    ball.x >= paddle.x &&
                    ball.x <= paddle.x + paddle.width) {

                    if (ball.dy > 0) {
                        ball.dy = -ball.dy;
                        // Add spin based on hit position
                        const hitPos = (ball.x - paddle.x) / paddle.width;
                        ball.dx = (hitPos - 0.5) * 8;
                    }
                }

                // Brick collision - improved detection with radius check
                const hasFire = gameState.activePowerups.fire;
                let brickHit = false;

                for (let j = gameState.bricks.length - 1; j >= 0; j--) {
                    const brick = gameState.bricks[j];

                    // Check collision with ball radius (more accurate)
                    const closestX = Math.max(brick.x, Math.min(ball.x, brick.x + brick.width));
                    const closestY = Math.max(brick.y, Math.min(ball.y, brick.y + brick.height));

                    const distanceX = ball.x - closestX;
                    const distanceY = ball.y - closestY;
                    const distanceSquared = (distanceX * distanceX) + (distanceY * distanceY);

                    if (distanceSquared < (ball.radius * ball.radius)) {
                        brickHit = true;

                        // Determine bounce direction based on where ball hit
                        const fromLeft = ball.x < brick.x;
                        const fromRight = ball.x > brick.x + brick.width;
                        const fromTop = ball.y < brick.y;
                        const fromBottom = ball.y > brick.y + brick.height;

                        // Bounce horizontally if hitting from sides
                        if (fromLeft || fromRight) {
                            ball.dx = -ball.dx;
                            // Push ball out of brick
                            if (fromLeft) ball.x = brick.x - ball.radius - 1;
                            if (fromRight) ball.x = brick.x + brick.width + ball.radius + 1;
                        }
                        // Bounce vertically if hitting from top/bottom
                        if (fromTop || fromBottom) {
                            ball.dy = -ball.dy;
                            // Push ball out of brick
                            if (fromTop) ball.y = brick.y - ball.radius - 1;
                            if (fromBottom) ball.y = brick.y + brick.height + ball.radius + 1;
                        }

                        // If corner hit, bounce both directions
                        if ((fromLeft || fromRight) && (fromTop || fromBottom)) {
                            ball.dx = -ball.dx;
                            ball.dy = -ball.dy;
                        }

                        // Damage brick
                        if (hasFire) {
                            brick.hits = 0;
                        } else {
                            brick.hits -= 1;
                        }

                        // Remove brick if destroyed
                        if (brick.hits <= 0) {
                            gameState.score += 10 * gameState.level;

                            // Spawn powerup if brick had one
                            if (brick.powerup) {
                                gameState.powerups.push({
                                    x: brick.x + brick.width / 2,
                                    y: brick.y,
                                    type: brick.powerup,
                                    width: 30,
                                    height: 30,
                                    dy: 2
                                });
                            }

                            gameState.bricks.splice(j, 1);
                        }

                        break; // Only hit one brick per frame
                    }
                }

                // Bottom collision (lose ball)
                if (ball.y - ball.radius > 600) {
                    gameState.balls.splice(i, 1);
                }
            }

            // If all balls lost, lose a life
            if (gameState.balls.length === 0) {
                gameState.lives -= 1;

                if (gameState.lives <= 0) {
                    gameState.gameOver = true;
                    saveHighScore();
                } else {
                    // Reset ball
                    const paddle = gameState.paddle;
                    gameState.balls = [{
                        x: paddle.x + paddle.width / 2,
                        y: paddle.y - 20,
                        dx: 3,
                        dy: -3,
                        radius: 8,
                        speed: 3
                    }];
                }
            }

            // Update powerups
            for (let i = gameState.powerups.length - 1; i >= 0; i--) {
                const powerup = gameState.powerups[i];
                powerup.y += powerup.dy;

                // Check paddle collision
                const paddle = gameState.paddle;
                if (powerup.y + powerup.height >= paddle.y &&
                    powerup.x + powerup.width >= paddle.x &&
                    powerup.x <= paddle.x + paddle.width) {

                    // Activate powerup
                    activatePowerup(powerup.type);
                    gameState.powerups.splice(i, 1);
                    gameState.score += 50;
                }
                // Remove if off screen
                else if (powerup.y > 600) {
                    gameState.powerups.splice(i, 1);
                }
            }

            // Check level complete
            if (gameState.bricks.length === 0) {
                gameState.levelComplete = true;
                gameState.score += 500 * gameState.level;
            }

            renderGame();
        }

        function activatePowerup(type) {
            gameState.activePowerups[type] = {
                active: true,
                startTime: Date.now()
            };

            // Apply immediate effects
            if (type === 'expand') {
                gameState.paddle.width = 150;
                setTimeout(function() { expirePowerup('expand'); }, 10000);
            } else if (type === 'slow') {
                for (let ball of gameState.balls) {
                    ball.dx *= 0.7;
                    ball.dy *= 0.7;
                }
                setTimeout(function() { expirePowerup('slow'); }, 10000);
            } else if (type === 'multi') {
                // Add two more balls
                const currentBalls = [...gameState.balls];
                for (let ball of currentBalls) {
                    for (let i = 0; i < 2; i++) {
                        const angle = Math.random() * Math.PI / 2 - Math.PI / 4;
                        const speed = Math.sqrt(ball.dx * ball.dx + ball.dy * ball.dy);
                        gameState.balls.push({
                            x: ball.x,
                            y: ball.y,
                            dx: speed * Math.sin(angle),
                            dy: -speed * Math.cos(angle),
                            radius: ball.radius,
                            speed: ball.speed
                        });
                    }
                }
            } else if (type === 'fire') {
                setTimeout(function() { expirePowerup('fire'); }, 10000);
            }
        }

        function expirePowerup(type) {
            if (!gameState || !gameState.activePowerups[type]) return;

            delete gameState.activePowerups[type];

            // Revert effects
            if (type === 'expand') {
                gameState.paddle.width = 100;
            } else if (type === 'slow') {
                for (let ball of gameState.balls) {
                    ball.dx /= 0.7;
                    ball.dy /= 0.7;
                }
            }
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
            gameLoop = setInterval(updateGame, 1000 / 60);  // 60 FPS
        }

        function startNewGame() {
            $('#gameOver').removeClass('show');
            $('#levelComplete').removeClass('show');
            isPaused = false;
            initGame(1);
        }

        function nextLevel() {
            $('#levelComplete').removeClass('show');
            isPaused = false;

            // Initialize next level
            const level = gameState.level + 1;
            const rows = Math.min(5 + level, 10);
            const colors = ['#ff0000', '#ff7700', '#ffff00', '#00ff00', '#0000ff', '#4b0082', '#9400d3'];

            const bricks = [];
            for (let row = 0; row < rows; row++) {
                for (let col = 0; col < 10; col++) {
                    const hasPowerup = Math.random() < 0.1;
                    const powerupType = hasPowerup ?
                        ['expand', 'slow', 'multi', 'fire'][Math.floor(Math.random() * 4)] : null;

                    bricks.push({
                        x: col * 60 + 5,
                        y: row * 20 + 50,
                        width: 55,
                        height: 15,
                        color: colors[row % colors.length],
                        hits: row < 3 ? 1 : (row < 6 ? 2 : 3),
                        maxHits: row < 3 ? 1 : (row < 6 ? 2 : 3),
                        powerup: powerupType
                    });
                }
            }

            // Update game state
            gameState.level = level;
            gameState.bricks = bricks;
            gameState.powerups = [];
            gameState.activePowerups = {};
            gameState.levelComplete = false;

            // Reset paddle and ball with increased speed
            gameState.paddle = {
                x: 250,
                y: 550,
                width: 100,
                height: 15,
                speed: 8
            };

            const baseSpeed = 3 + level * 0.5;
            gameState.balls = [{
                x: 300,
                y: 535,
                dx: baseSpeed,
                dy: -baseSpeed,
                radius: 8,
                speed: baseSpeed
            }];

            renderGame();
            startGameLoop();
        }

        function togglePause() {
            if (gameState && !gameState.gameOver && !gameState.levelComplete) {
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
                            const date = score.date ? ' (' + score.date.split(' ')[0] + ')' : '';

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

        // Mouse and touch tracking
        function updateMousePosition(clientX) {
            const rect = canvas.getBoundingClientRect();
            const scaleX = canvas.width / rect.width;
            mouseX = Math.max(0, Math.min(600, (clientX - rect.left) * scaleX));
        }

        canvas.addEventListener('mousemove', function(e) {
            e.preventDefault();
            updateMousePosition(e.clientX);
        });

        canvas.addEventListener('touchmove', function(e) {
            e.preventDefault();
            if (e.touches.length > 0) {
                updateMousePosition(e.touches[0].clientX);
            }
        });

        canvas.addEventListener('touchstart', function(e) {
            e.preventDefault();
            if (e.touches.length > 0) {
                updateMousePosition(e.touches[0].clientX);
            }
        });

        // Also track on game board container
        document.getElementById('game-board').addEventListener('mousemove', function(e) {
            updateMousePosition(e.clientX);
        });

        // Keyboard controls
        $(document).keydown(function(e) {
            if (e.keyCode === 80) {  // P for pause
                togglePause();
                e.preventDefault();
            }
        });

        // Start game on load
        $(document).ready(function() {
            console.log('Document ready, initializing game...');
            try {
                initGame(1);
                console.log('Game initialized successfully');
            } catch(e) {
                console.error('Error initializing game:', e);
                alert('Error loading game: ' + e.message);
            }
        });
    </script>
</body>
</html>
'''
