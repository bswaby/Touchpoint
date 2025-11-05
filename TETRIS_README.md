# TouchPoint Tetris Game

## Installation Instructions

### 1. Add Python Script to TouchPoint

1. Log into your TouchPoint admin panel
2. Navigate to **Admin → Advanced → Special Content → Python**
3. Click **"Add New"**
4. Name it: `TetrisGame`
5. Copy and paste the contents of `TetrisGame.py` into the script editor
6. Click **Save**

### 2. Add to CustomReports (Optional)

If you want the game accessible from the Blue Toolbar or reports menu:

1. Navigate to **Special Content → Text Content → CustomReports**
2. Add this line:
```xml
<Report name="TetrisGame" type="PyScript" role="Access" />
```
3. Save (note: may take up to 24 hours to appear due to caching)

### 3. Access the Game

Navigate to:
```
https://your-touchpoint-domain.com/PyScriptForm/TetrisGame
```

**Note:** Use `/PyScriptForm/` (not `/PyScript/`) to render the HTML game interface!

## How to Play

### Controls

| Key | Action |
|-----|--------|
| **←** | Move piece left |
| **→** | Move piece right |
| **↓** | Soft drop (move down faster) |
| **↑** | Rotate piece clockwise |
| **Space** | Hard drop (instant drop) |
| **P** | Pause/Resume game |

### Objective

- Arrange falling tetrominoes to create complete horizontal lines
- Complete lines disappear and award points
- Game ends when pieces stack to the top
- Survive as long as possible and achieve the highest score!

### Scoring

- **Single line:** 100 × Level
- **Double line:** 300 × Level
- **Triple line:** 500 × Level
- **Tetris (4 lines):** 800 × Level
- **Soft drop:** 1 point per cell
- **Hard drop:** 2 points per cell

### Leveling

- Start at Level 1
- Level up every 10 lines cleared
- Speed increases with each level
- Maximum challenge at higher levels!

## Features

### High Score System
- Automatically saves your top scores
- Leaderboard displays top 10 players
- Persistent storage using TouchPoint content
- Enter your name after game over

### Next Piece Preview
- See the upcoming piece
- Plan your strategy ahead
- Displayed in the sidebar

### Statistics
- Real-time score tracking
- Current level display
- Total lines cleared

## Technical Details

### Storage
The game uses TouchPoint's content storage system:
- High scores stored in: `TetrisHighScores` (Special Content → Text)
- Format: JSON array of score objects
- Maximum 10 scores retained

### Compatibility
- Python: IronPython 2.7.3
- Browser: Any modern browser with HTML5 Canvas support
- Mobile: Touch controls not implemented (keyboard only)

### Customization

You can modify these constants in the JavaScript section:

```javascript
const COLS = 10;        // Board width
const ROWS = 20;        // Board height
const BLOCK_SIZE = 30;  // Pixel size of each block
```

Colors can be customized in the `COLORS` array.

## Troubleshooting

### Game doesn't load
- Ensure you're using `/PyScriptForm/` not `/PyScript/`
- Check browser console for JavaScript errors
- Verify script is saved in TouchPoint Python content

### High scores not saving
- Check TouchPoint permissions for content storage
- Verify AJAX requests aren't being blocked
- Check browser network tab for failed POST requests

### Controls not responding
- Click on the game canvas to ensure it has focus
- Check that no browser extensions are intercepting key presses
- Try refreshing the page

## Credits

Classic Tetris gameplay implemented for TouchPoint using:
- IronPython 2.7.3 backend
- HTML5 Canvas rendering
- Vanilla JavaScript game loop
- TouchPoint content storage API

Enjoy the game! 🎮

## Future Enhancements (Ideas)

- Mobile touch controls
- Sound effects and background music
- Ghost piece (preview of drop location)
- Hold piece functionality
- Multiplayer mode
- Daily challenges
- Achievement system
