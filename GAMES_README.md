# TouchPoint Retro Games Collection

Two classic arcade games built with IronPython 2.7.3 for TouchPoint!

## 🎮 Games Included

### 1. Tetris - Classic Block Puzzle
**File:** `TetrisGame.py`
**URL:** `/PyScriptForm/TetrisGame`

The timeless puzzle game where you arrange falling blocks to clear lines.

**Quick Stats:**
- 7 tetromino shapes
- Progressive difficulty
- Level up every 10 lines
- High score leaderboard
- Smooth rotation & controls

**Perfect for:** Strategy lovers, puzzle enthusiasts

---

### 2. Breakout - Brick Breaker
**File:** `BreakoutGame.py`
**URL:** `/PyScriptForm/BreakoutGame`

Break all the bricks with your paddle and ball. Catch power-ups for special abilities!

**Quick Stats:**
- 4 unique power-ups
- Multiple lives system
- Progressive levels
- High score leaderboard
- Mouse + keyboard controls

**Perfect for:** Action fans, arcade enthusiasts

---

## 🚀 Quick Installation (Both Games)

### Step 1: Upload Scripts
1. Go to **Admin → Advanced → Special Content → Python**
2. Click **"Add New"** for each game:
   - Add `TetrisGame` (paste TetrisGame.py content)
   - Add `BreakoutGame` (paste BreakoutGame.py content)
3. Save each script

### Step 2: Access Games
Navigate to:
- **Tetris:** `https://your-domain.com/PyScriptForm/TetrisGame`
- **Breakout:** `https://your-domain.com/PyScriptForm/BreakoutGame`

### Step 3 (Optional): Add to Menu
Edit **Special Content → Text Content → CustomReports**:
```xml
<Report name="TetrisGame" type="PyScript" role="Access" />
<Report name="BreakoutGame" type="PyScript" role="Access" />
```

---

## 🎯 Controls Comparison

| Action | Tetris | Breakout |
|--------|--------|----------|
| **Move Left** | ← | ← or Mouse |
| **Move Right** | → | → or Mouse |
| **Special Action** | ↑ Rotate | Space Launch |
| **Quick Drop** | ↓ Soft Drop | - |
| **Instant Drop** | Space | - |
| **Pause** | P | P |

---

## 🏆 Features Comparison

| Feature | Tetris | Breakout |
|---------|--------|----------|
| **Difficulty** | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| **Skill Ceiling** | Very High | High |
| **Game Length** | 5-20 min | 10-30 min |
| **Lives System** | ❌ | ✅ (3 lives) |
| **Power-ups** | ❌ | ✅ (4 types) |
| **Levels** | Auto-progression | Manual advancement |
| **High Scores** | ✅ Top 10 | ✅ Top 10 |
| **Mouse Control** | ❌ | ✅ |

---

## 💾 Data Storage

Both games store high scores in TouchPoint's Special Content:
- **Tetris:** `TetrisHighScores`
- **Breakout:** `BreakoutHighScores`

Format: JSON array, top 10 scores only

To view/manage:
1. Go to **Special Content → Text Content**
2. Search for game name
3. View or edit JSON data

To reset leaderboard:
- Delete the content or replace with `[]`

---

## 🎨 Technical Implementation

### Architecture
```
┌─────────────────────────────────────┐
│   IronPython 2.7.3 Backend          │
│   • High score management           │
│   • Leaderboard persistence         │
│   • HTML/JS generation              │
└─────────────────────────────────────┘
                 ↓
┌─────────────────────────────────────┐
│   HTML5 Canvas Frontend             │
│   • Real-time rendering (60 FPS)    │
│   • Game physics & collision        │
│   • Input handling                  │
└─────────────────────────────────────┘
                 ↓
┌─────────────────────────────────────┐
│   AJAX Communication                │
│   • Score submission                │
│   • Leaderboard updates             │
│   • No page refresh needed          │
└─────────────────────────────────────┘
```

### Why This Design?

**Server-side (Python):**
- Persistent data storage
- High score validation
- Leaderboard management

**Client-side (JavaScript):**
- Real-time gameplay (no latency)
- Smooth 60 FPS rendering
- Instant input response

**Result:** Best of both worlds!

---

## 📊 Scoring Systems

### Tetris Scoring
```
Single Line:  100 × Level
Double Line:  300 × Level
Triple Line:  500 × Level
Tetris (4):   800 × Level
Soft Drop:    1 point/cell
Hard Drop:    2 points/cell

Level Up: Every 10 lines
Speed: +100ms faster per level
```

### Breakout Scoring
```
Brick Values (× Level):
  Red:    50 pts
  Orange: 40 pts
  Yellow: 30 pts
  Green:  20 pts
  Blue:   10 pts

Level Complete Bonus:
  Lives × 1000
  Remaining Bricks × 50
```

---

## 🔧 Customization Guide

### Modify Tetris Difficulty
```javascript
// In TetrisGame.py, find these lines:
const COLS = 10;        // Change board width
const ROWS = 20;        // Change board height
const BLOCK_SIZE = 30;  // Change piece size

// Change starting speed
dropInterval = 1000;    // Milliseconds (lower = faster)

// Change level progression
level = Math.floor(lines / 10) + 1;  // Change 10 to adjust
```

### Modify Breakout Difficulty
```javascript
// In BreakoutGame.py, find these lines:
const PADDLE_WIDTH = 100;   // Change paddle size
const BRICK_ROWS = 5;       // More rows = harder
const BRICK_COLS = 10;      // More columns = longer

// Starting lives
let lives = 3;              // Change starting lives

// Power-up drop rate
if (Math.random() < 0.1)    // Change 0.1 (10% chance)
```

### Change Colors
Both games have color arrays you can modify:
```javascript
// Tetris
const COLORS = [
    null,
    '#FF0D72',  // I piece - Modify these hex codes
    '#0DC2FF',  // O piece
    // ... etc
];

// Breakout
const BRICK_COLORS = [
    '#FF0D72',  // Top row
    '#FF8E0D',  // Row 2
    // ... etc
];
```

---

## 🐛 Common Issues & Solutions

### Issue: Games don't load
**Solution:**
- Use `/PyScriptForm/` NOT `/PyScript/`
- Check browser console (F12) for errors
- Verify script names match exactly

### Issue: Controls don't work
**Solution:**
- Click on the game canvas first
- Check keyboard isn't being captured by other apps
- Try refreshing the page (F5)

### Issue: High scores won't save
**Solution:**
- Check TouchPoint content permissions
- View browser Network tab (F12) for failed requests
- Verify AJAX isn't blocked by firewall

### Issue: Laggy performance
**Solution:**
- Close other browser tabs
- Disable browser extensions
- Update browser to latest version
- Check computer performance (CPU/RAM)

### Issue: Games too easy/hard
**Solution:**
- See Customization Guide above
- Adjust speed, size, or difficulty constants
- Modify scoring multipliers

---

## 🎓 Game Mechanics Deep Dive

### Tetris Strategy Guide
1. **Keep it flat:** Avoid spikes and valleys
2. **Save the well:** Keep one column for I-pieces
3. **Think ahead:** Use the next piece preview
4. **T-spins:** Advanced rotation techniques
5. **Speed management:** Don't panic at high levels

### Breakout Strategy Guide
1. **Edge control:** Master paddle edge shots for angles
2. **Power-up priority:** Expand > Slow > Multi > Fire
3. **Top-down:** Break top rows first (more points)
4. **Channel building:** Create vertical paths
5. **Life conservation:** Play safe on last life

---

## 📱 Browser Compatibility

| Browser | Tetris | Breakout |
|---------|--------|----------|
| Chrome | ✅ Perfect | ✅ Perfect |
| Firefox | ✅ Perfect | ✅ Perfect |
| Safari | ✅ Perfect | ✅ Perfect |
| Edge | ✅ Perfect | ✅ Perfect |
| Mobile Safari | ⚠️ Keyboard only | ⚠️ Keyboard only |
| Mobile Chrome | ⚠️ Keyboard only | ⚠️ Keyboard only |

**Note:** Mobile support is limited due to keyboard-focused controls. Future versions may add touch support.

---

## 🎉 Leaderboard Competition Ideas

### Office Tournament
1. Set up public access URL
2. Announce competition period
3. Top 3 players win prizes
4. Reset leaderboard monthly

### Department Challenges
1. Create separate game instances per department
2. Compare average scores
3. Winning department gets bragging rights
4. Rotate game monthly (Tetris ↔ Breakout)

### Break Room Fun
1. Display game URL on screen in break room
2. Rotate leaderboard display every 30 seconds
3. Update weekly champions board
4. Create friendly competition

---

## 📝 Credits & License

**Created for TouchPoint using:**
- IronPython 2.7.3
- HTML5 Canvas API
- Vanilla JavaScript (no frameworks)
- TouchPoint Content Storage API

**Game Inspirations:**
- Tetris: Original by Alexey Pajitnov (1984)
- Breakout: Original by Atari (1976)

**Implementation:**
- Modern web technologies
- Responsive design
- High score persistence
- Mobile-friendly UI (keyboard required)

---

## 🚀 What's Next?

### Potential Future Games
- **Snake** - Classic Nokia game
- **Space Invaders** - Shoot the aliens
- **Pac-Man** - Maze chase game
- **Pong** - Two-player classic
- **Asteroids** - Space shooter

### Feature Requests
Want a new feature? Consider:
- Sound effects and music
- Mobile touch controls
- Achievements and badges
- Daily challenges
- Multiplayer modes
- Custom skins/themes

---

## 💬 Support

### Getting Help
1. Check the individual game README files
2. Review troubleshooting sections
3. Test in different browsers
4. Check TouchPoint permissions

### Reporting Issues
When reporting problems, include:
- Browser and version
- TouchPoint version
- Steps to reproduce
- Console errors (F12 → Console)
- Network errors (F12 → Network)

---

## 🎮 Have Fun!

Enjoy your retro gaming break in TouchPoint! May your Tetris lines be straight and your Breakout paddle never miss!

**Quick Links:**
- [Tetris README](TETRIS_README.md)
- [Breakout README](BREAKOUT_README.md)

---

*Built with ❤️ for the TouchPoint community*
