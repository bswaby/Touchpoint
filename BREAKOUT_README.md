# TouchPoint Breakout Game

## Installation Instructions

### 1. Add Python Script to TouchPoint

1. Log into your TouchPoint admin panel
2. Navigate to **Admin → Advanced → Special Content → Python**
3. Click **"Add New"**
4. Name it: `BreakoutGame`
5. Copy and paste the contents of `BreakoutGame.py` into the script editor
6. Click **Save**

### 2. Add to CustomReports (Optional)

If you want the game accessible from the Blue Toolbar or reports menu:

1. Navigate to **Special Content → Text Content → CustomReports**
2. Add this line:
```xml
<Report name="BreakoutGame" type="PyScript" role="Access" />
```
3. Save (note: may take up to 24 hours to appear due to caching)

### 3. Access the Game

Navigate to:
```
https://your-touchpoint-domain.com/PyScriptForm/BreakoutGame
```

**Note:** Use `/PyScriptForm/` (not `/PyScript/`) to render the HTML game interface!

## How to Play

### Controls

| Control | Action |
|---------|--------|
| **Mouse** | Move paddle left/right |
| **←** | Move paddle left (keyboard) |
| **→** | Move paddle right (keyboard) |
| **Space** | Launch ball |
| **P** | Pause/Resume game |

### Objective

- Use the paddle to bounce the ball and destroy all bricks
- Different colored bricks are worth different points
- Catch power-ups for special abilities
- Complete all levels with the highest score!
- Don't let the ball fall off the bottom of the screen

### Scoring

**Brick Values (multiplied by level):**
- 🔴 Red (Top): 50 points
- 🟠 Orange: 40 points
- 🟡 Yellow: 30 points
- 🟢 Green: 20 points
- 🔵 Blue (Bottom): 10 points

**Bonuses:**
- Level completion: Lives × 1000 points
- Remaining bricks: 50 points each

### Lives System

- Start with 3 lives
- Lose a life when ball falls off screen
- Game over when all lives are lost
- Extra lives awarded for level completion bonuses

## Power-ups

### 🟢 Expand Paddle
- **Effect:** Paddle becomes 50% wider
- **Duration:** 10 seconds
- **Strategy:** Easier to catch the ball and hit bricks

### 🔵 Slow Ball
- **Effect:** Ball moves 30% slower
- **Duration:** 8 seconds
- **Strategy:** Better control and precision aiming

### 🟡 Multi-Ball
- **Effect:** Spawns 2 additional balls
- **Duration:** Until balls are lost
- **Strategy:** Hit multiple bricks simultaneously

### 🔴 Fire Ball
- **Effect:** Ball pierces through multiple bricks
- **Duration:** 10 seconds
- **Strategy:** Massive brick destruction, especially effective on dense areas

## Level Progression

### Level Difficulty
- **Level 1:** Single-hit bricks
- **Level 2+:** Bricks require multiple hits (up to 3)
- **Ball Speed:** Increases with each level
- **Challenge:** Progressive difficulty ramp

### Brick Durability
- Bricks fade as they take damage
- Higher levels = more hits required
- Fire ball ignores brick durability

### Level Completion
- Destroy all bricks to complete level
- Receive completion bonus
- Automatically advance to next level
- Paddle resets, ball speed increases

## Strategies & Tips

### Paddle Control
1. **Mouse is best:** Most precise control
2. **Keyboard works:** Use arrow keys if needed
3. **Edge shots:** Hit ball on paddle edge for sharp angles
4. **Center shots:** Hit center for vertical trajectory

### Brick Breaking
1. **Top-down approach:** Higher bricks = more points
2. **Create channels:** Break vertical paths through bricks
3. **Side walls:** Use walls to reach difficult angles
4. **Fire ball timing:** Save for dense brick clusters

### Power-up Management
1. **Prioritize Expand:** Safest, most forgiving
2. **Slow ball first:** Easier to catch other power-ups
3. **Multi-ball carefully:** Can be chaotic, but powerful
4. **Fire ball strategically:** Best for clearing tough spots

### Advanced Techniques
1. **Angle control:** Position paddle to direct ball
2. **Bounce prediction:** Anticipate ball trajectory
3. **Power-up stacking:** Combine effects when possible
4. **Life conservation:** Play safe on last life

## Technical Details

### Storage
The game uses TouchPoint's content storage system:
- High scores stored in: `BreakoutHighScores` (Special Content → Text)
- Format: JSON array of score objects
- Maximum 10 scores retained

### Physics
- Ball speed: Increases with level (3 + level × 0.5)
- Paddle angle: Hit position affects ball direction
- Collision detection: Pixel-perfect hitboxes
- Frame rate: 60 FPS via requestAnimationFrame

### Compatibility
- Python: IronPython 2.7.3
- Browser: Any modern browser with HTML5 Canvas support
- Canvas: 600×600 pixels
- Mouse: Required for optimal gameplay

### Customization

You can modify these constants in the JavaScript section:

```javascript
const PADDLE_WIDTH = 100;    // Paddle width in pixels
const BRICK_ROWS = 5;        // Number of brick rows
const BRICK_COLS = 10;       // Number of brick columns
const BALL_RADIUS = 8;       // Ball size

// Power-up drop chance (0.0 to 1.0)
if (Math.random() < 0.1) // 10% chance
```

Brick colors and values:
```javascript
const BRICK_COLORS = ['#FF0D72', '#FF8E0D', '#FFE138', '#0DFF72', '#0DC2FF'];
const BRICK_VALUES = [50, 40, 30, 20, 10];
```

## Troubleshooting

### Game doesn't load
- Ensure you're using `/PyScriptForm/` not `/PyScript/`
- Check browser console for JavaScript errors
- Verify script is saved in TouchPoint Python content

### Paddle doesn't move
- Make sure mouse is over the canvas
- Try keyboard controls (arrow keys)
- Check browser mouse permissions

### High scores not saving
- Check TouchPoint permissions for content storage
- Verify AJAX requests aren't being blocked
- Check browser network tab for failed POST requests

### Ball goes through bricks
- Normal behavior with Fire Ball power-up
- Otherwise, may indicate browser performance issue
- Try closing other tabs to free up resources

### Performance issues
- Close unnecessary browser tabs
- Disable browser extensions temporarily
- Reduce canvas quality in browser settings

## Credits

Classic Breakout/Arkanoid gameplay implemented for TouchPoint using:
- IronPython 2.7.3 backend
- HTML5 Canvas rendering
- Vanilla JavaScript game loop
- TouchPoint content storage API
- Mouse and keyboard input handling

Enjoy breaking bricks! 🎯

## Future Enhancements (Ideas)

- Sound effects and music
- Touch controls for mobile
- Boss levels with moving bricks
- Combo multipliers
- Achievement system
- Replay/ghost mode
- Custom level designer
- Particle effects on brick destruction
