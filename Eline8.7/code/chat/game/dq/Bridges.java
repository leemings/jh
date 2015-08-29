/* Bridges Source Code          */
/* By Dick Christoph April 1998 */
/* dchristo@minn.net            */
/* www1.minn.net/~dchristo      */

import java.applet.*;
import java.awt.*;
import java.io.*;

public class Bridges extends Applet {
  public static final int MAXSCREENX = 340;
  public static final int MAXSCREENY = 340;
  public static final int MAXROWS = 17;
  public static final int MAXCOLS = 17;
  public static final int NOTOWER = 0;
  public static final int BLUETOWER = 1, BLUE = 1;
  public static final int REDTOWER = 2, RED = 2;
  public static final int ROW = 0;
  public static final int COL = 1;

  public static final int UP = 1;
  public static final int RIGHT = 2;
  public static final int LEFT = 3;
  public static final int DOWN = 4;

  public static int board[][], whoseTurn;
  static BridgeCanvas cnv1;
  static Label turnMsg, statusMsg;
  Button newGameButton;
  Panel p, p1;

  static boolean redWins, blueWins, debug;

  public static final int MAXSTEPS = 10000;
  public static final int MAXOPTIONS = 200;
  public static final int MAXPATHLIST = 32;
  public static boolean visited[][];
  public static boolean consider[][];
  public static Option optionList[];
  public static Step steps[];
  public static int minSteps[][];
  public static PathList bluePathList[], redPathList[];
  public static int optionTop, firstFreeStep, bluePathListTop, redPathListTop;
  public static double stepCounter;
  public static double maxSteps, maxOptions, maxPathList, maxMoves;

  public void init()
  { 
    setLayout(null);
    resize(MAXSCREENX,MAXSCREENY+60);
    addNotify();

    debug = false;
    cnv1 = new BridgeCanvas();
    add(cnv1);
    cnv1.reshape(0,0,MAXSCREENX, MAXSCREENY);

    p = new Panel();
    turnMsg = new Label("");
    turnMsg.setAlignment(Label.CENTER);
    p.add(turnMsg);
    whoseTurn = BLUE;
    toggleTurnMsg();
    p.reshape(0,MAXSCREENY,MAXSCREENX,30);
    add(p);
    p1 = new Panel();
    newGameButton = new Button("New Game");
    p1.add(newGameButton);
    if (debug)
    { statusMsg = new Label("");
      statusMsg.setAlignment(Label.CENTER);
      p1.add(statusMsg);
    }
    p1.reshape(0,MAXSCREENY+30,MAXSCREENX,30);
    add(p1);
    initBoard();
    setupMem();
    cnv1.repaint();
  }

  public boolean action(Event e, Object arg)
  { Object source = e.target;
    if (source == newGameButton)
    { initBoard();
      cnv1.repaint();
    }
    return true;
  }

  public void destroy() {
  }

  public void initBoard()
  { int r, c;
    redWins = false;
    blueWins = false;
    whoseTurn = BLUE;
    toggleTurnMsg();
    if (debug)
      statusMsg.setText("");
    maxSteps = 0;
    maxOptions = 0;
    maxPathList = 0;
    maxMoves = 0;

    board = new int[MAXROWS][MAXCOLS];

    for (r = 0; r < MAXROWS; r++)
    { for (c = 0; c < MAXCOLS; c++)
      { board[r][c] = NOTOWER;
        if ((r % 2) == 0)   /* even row */
        { if ((c % 2) == 1)   /* odd col */
          { board[r][c] = BLUETOWER;
        } }
        else  /* odd row */
        { if ((c % 2) == 0)   /* even col */
          { board[r][c] = REDTOWER;
    } } } }
  }

  public static void toggleTurnMsg()
  {
    if (whoseTurn == BLUE)
    { whoseTurn = RED;
      turnMsg.setText("Red - It is your Move");
    }
    else
    { whoseTurn = BLUE;
      turnMsg.setText("Blue - It is your Move");
    }
  }

  public static boolean toggleBox(int r, int c)
  { boolean rv;

    if (board[r][c] == NOTOWER)
    { rv = true;
      if (whoseTurn == BLUE)
      { board[r][c] = BLUETOWER;
        toggleTurnMsg();
        determineBlueProgress();
      }
      else
      { board[r][c] = REDTOWER;
        toggleTurnMsg();
        determineRedProgress();
      }
    }
    else
      rv = false;
    return(rv);
  }

  public static void determineBlueProgress()
  { int r,c, r1, c1, tryTop, tryPtr;
    int tries[][];

    visited = new boolean[MAXROWS][MAXCOLS];
    for (r = 0; r < MAXROWS; r++)
      for (c = 0; c < MAXCOLS; c++)
        visited[r][c] = false;

    tries = new int[MAXROWS*MAXCOLS][2];
    blueWins = false;

    r = 0;
    for (c = 1; c < MAXCOLS && (!blueWins); c = c + 2)
    { /* board[0][c] = BLUETOWER */
      if (!visited[0][c])
      { tryTop = 0;
        tryPtr = 0;
        tries[tryTop][ROW] = 0;
        tries[tryTop++][COL] = c;
        while ((tryPtr < tryTop) && (!blueWins))
        { r1 = tries[tryPtr][ROW];
          c1 = tries[tryPtr][COL];
          visited[r1][c1] = true;
          if (r1 == MAXROWS - 1)
            blueWins = true;
          else
          { if ((r1 - 1) >= 0)  /* check up */
            { if ((!visited[r1-1][c1]) && (board[r1-1][c1] == BLUETOWER))
              { tries[tryTop][ROW] = r1-1;
                tries[tryTop++][COL] = c1;
            } }
            /* check down */
            if ((!visited[r1+1][c1]) && (board[r1+1][c1] == BLUETOWER))
            { { tries[tryTop][ROW] = r1+1;
                tries[tryTop++][COL] = c1;
            } }
            if ((c1 - 1) >= 0)  /* check left */
            { if ((!visited[r1][c1-1]) && (board[r1][c1-1] == BLUETOWER))
              { tries[tryTop][ROW] = r1;
                tries[tryTop++][COL] = c1-1;
            } }
            /* check right */
            if ((c1 + 1) < MAXCOLS)  
            { if ((!visited[r1][c1+1]) && (board[r1][c1+1] == BLUETOWER))
              { tries[tryTop][ROW] = r1;
                tries[tryTop++][COL] = c1+1;
          } } }
          tryPtr++;
        }
      }
    }
    if (blueWins)
    { turnMsg.setText("Blue Wins");
      if (debug)
        statusMsg.setText("Blue Wins" + " MS: " + maxSteps + " MO: " + maxOptions + " MP: " + maxPathList + " MM: " + maxMoves);
    }
  }

  public static void determineRedProgress()
  { int r,c, r1, c1, tryTop, tryPtr;
    int tries[][];

    visited = new boolean[MAXROWS][MAXCOLS];
    for (r = 0; r < MAXROWS; r++)
      for (c = 0; c < MAXCOLS; c++)
        visited[r][c] = false;

    tries = new int[MAXROWS*MAXCOLS][2];
    redWins = false;

    c = 0;
    for (r = 1; r < MAXROWS && (!redWins); r = r + 2)
    { /* board[r][0] = REDTOWER */
      if (!visited[r][0])
      { tryTop = 0;
        tryPtr = 0;
        tries[tryTop][ROW] = r;
        tries[tryTop++][COL] = 0;
        while ((tryPtr < tryTop) && (!redWins))
        { r1 = tries[tryPtr][ROW];
          c1 = tries[tryPtr][COL];
          visited[r1][c1] = true;
          if (c1 == MAXCOLS - 1)
            redWins = true;
          else
          { if ((r1 - 1) >= 0)  /* check up */
            { if ((!visited[r1-1][c1]) && (board[r1-1][c1] == REDTOWER))
              { tries[tryTop][ROW] = r1-1;
                tries[tryTop++][COL] = c1;
            } }
            if ((r1 + 1) < MAXROWS)  /* check down */
            { if ((!visited[r1+1][c1]) && (board[r1+1][c1] == REDTOWER))
              { tries[tryTop][ROW] = r1+1;
                tries[tryTop++][COL] = c1;
            } }
            if ((c1 - 1) >= 0)  /* check left */
            { if ((!visited[r1][c1-1]) && (board[r1][c1-1] == REDTOWER))
              { tries[tryTop][ROW] = r1;
                tries[tryTop++][COL] = c1-1;
            } }
            /* check right */
            if ((c1 + 1) < MAXCOLS)  
            { if ((!visited[r1][c1+1]) && (board[r1][c1+1] == REDTOWER))
              { tries[tryTop][ROW] = r1;
                tries[tryTop++][COL] = c1+1;
          } } }
          tryPtr++;
        }
      }
    }
    if (redWins)
    { turnMsg.setText("Red Wins");
      if (debug)
        statusMsg.setText("Red Wins" + " MS: " + maxSteps + " MO: " + maxOptions + " MP: " + maxPathList + " MM: " + maxMoves);
  } }

  public static void setupMem()
  { int i;

    steps = new Step[MAXSTEPS];
    for (i = 0; i < MAXSTEPS; i++)
      steps[i] = new Step();

    optionList = new Option[MAXOPTIONS];
    for (i = 0; i < MAXOPTIONS; i++)
      optionList[i] = new Option();

    bluePathList = new PathList[MAXPATHLIST];
    for (i = 0; i < MAXPATHLIST; i++)
      bluePathList[i] = new PathList();

    redPathList = new PathList[MAXPATHLIST];
    for (i = 0; i < MAXPATHLIST; i++)
      redPathList[i] = new PathList();
  }

  public static void findBestMove()
  { int i, r, c, pth, stp, s, curStep, row, col;
    Point bestMove;

    stepCounter = 0;
    for (i = 0; i < MAXSTEPS; i++)
      steps[i].nextStep = i + 1;
    steps[MAXSTEPS-1].nextStep = -1;
    firstFreeStep = 0;

    visited = new boolean[MAXROWS][MAXCOLS];
    minSteps = new int[MAXROWS][MAXCOLS];
    consider = new boolean[MAXROWS][MAXCOLS];
    for (r = 0; r < MAXROWS; r++)
    { for (c = 0; c < MAXCOLS; c++)
      { visited[r][c] = false;
        minSteps[r][c] = MAXROWS * MAXCOLS;
        if ((r == 0) || (r == MAXROWS-1) || (c == 0) || (c == MAXCOLS-1))
          consider[r][c] = false;
        else
          consider[r][c] = true;
      }
    }

    optionTop = 0;
    bluePathListTop = 0;
    findShortestBluePaths();

    for (r = 0; r < MAXROWS; r++)
    { for (c = 0; c < MAXCOLS; c++)
      { visited[r][c] = false;
        minSteps[r][c] = MAXROWS * MAXCOLS;
        if ((r == 0) || (r == MAXROWS-1) || (c == 0) || (c == MAXCOLS-1))
          consider[r][c] = false;
        else
          consider[r][c] = true;
      }
    }
    optionTop = 0;
    redPathListTop = 0;
    findShortestRedPaths();
    if (debug)
      statusMsg.setText("BPLT: " + bluePathListTop + " RPLT: " + redPathListTop + " SC: " + stepCounter + " MS: " + maxSteps);
    bestMove = compareBlueAndRedLists();
    row = bestMove.y;
    col = bestMove.x;
    toggleBox(row,col);
    cnv1.repaint();
  }

  public static final int COUNT = 2, CLOSEST = 3;
  public static final int MAXDIST = MAXROWS * MAXCOLS;

  public static Point compareBlueAndRedLists()
  { int blueList[][], redList[][], tempBoard[][][], moveList[][];
    int blueTop, redTop, cmp, b, r, c, p, row, col, cp, lp, i, j;
    int close, closeCnt, lo, hi, mid, maxCount, count, moveListTop;

    boolean found, secureTop, secureBottom;
    Point rv;

    rv = new Point(0,0);

    secureTop = false;
    secureBottom = false;
    for (c = 1; c < MAXCOLS; c+=2)
    {  if (board[1][c] == BLUETOWER)
        secureTop = true;
      if (board[MAXROWS-2][c] == BLUETOWER)
        secureBottom = true;
    }
    if ((!secureTop) || (!secureBottom))
    { if (!secureTop)
      { for (b = 0; (b <= 6) && (!secureTop); b += 2)
        { for (c = 7-b; c <= 9+b; c += 2)
          { if (board[1][c] == NOTOWER)
            { rv.x = c;
              rv.y = 1;
              secureTop = true;
      } } } }
      else
      { for (b = 0; (b <= 6) && (!secureBottom); b += 2)
        { for (c = 7-b; c <= 9+b; c += 2)
          { if (board[MAXROWS-2][c] == NOTOWER)
            { rv.x = c;
              rv.y = MAXROWS-2;
              secureBottom = true;
    } } } } }
    else
    { 
      if ((redPathList[0].togo == 1) && (bluePathList[0].togo > 1))
      { /* Red has only one togo, must block it */
        cp = redPathList[0].firstStep;
        rv.x = -1;
        while ((cp != -1) && (rv.x == -1))
        { row = steps[cp].row;
          col = steps[cp].col;
          if (board[row][col] == NOTOWER)
          { rv.x = col;
            rv.y = row;
          }
          cp = steps[cp].nextStep;
        }
      }
      else
      { blueList = new int[MAXROWS*MAXCOLS][3];
        redList = new int[MAXROWS*MAXCOLS][4];
        tempBoard = new int[MAXROWS][MAXCOLS][2];  /* 0 = COUNT 1 = CLOSEST */
        moveList = new int[MAXROWS*MAXCOLS][2];
  
        for (r = 0; r < MAXROWS; r++)
        { for (c = 0; c < MAXCOLS; c++)
          { tempBoard[r][c][0] = 0;
            tempBoard[r][c][1] = 0;
        } }
  
        for (p = 0; p < bluePathListTop; p++)
        { cp = bluePathList[p].firstStep;
          while (cp != -1)
          { row = steps[cp].row;
            col = steps[cp].col;
            if (board[row][col] == NOTOWER)
            { tempBoard[row][col][0] += 1;
            }
            cp = steps[cp].nextStep;
        } }

        blueTop = 0;
        for (r = 0; r < MAXROWS; r++)
        { for (c = 0; c < MAXCOLS; c++)
          { if (tempBoard[r][c][0] != 0)
            { blueList[blueTop][ROW] = r;
              blueList[blueTop][COL] = c;
              blueList[blueTop++][COUNT] = tempBoard[r][c][0];
        } } }

        for (r = 0; r < MAXROWS; r++)
        { for (c = 0; c < MAXCOLS; c++)
          { tempBoard[r][c][0] = 0;
            tempBoard[r][c][1] = MAXDIST;
        } }
  
        lp = -1;
        for (p = 0; p < redPathListTop; p++)
        { cp = redPathList[p].firstStep;
          steps[cp].prevStep = -1;    /* This should have been taken care of elsewhere */
          closeCnt = MAXDIST;
          while (cp != -1)
          { row = steps[cp].row;
            col = steps[cp].col;
            if (board[row][col] == NOTOWER)
            { tempBoard[row][col][0] += 1;
              tempBoard[row][col][1] = Math.min(tempBoard[row][col][1], closeCnt++);
            }
            else
              closeCnt = 1;
            lp = cp;
            cp = steps[cp].nextStep;
          }
          closeCnt = MAXDIST;
          while (lp != -1)
          { row = steps[lp].row;
            col = steps[lp].col;
            if (board[row][col] == NOTOWER)
            { tempBoard[row][col][0] += 1;
              tempBoard[row][col][1] = Math.min(tempBoard[row][col][1], closeCnt++);
            }
            else
              closeCnt = 1;
            lp = steps[lp].prevStep;
          }
        }
  
        redTop = 0;
        for (r = 0; r < MAXROWS; r++)
        { for (c = 0; c < MAXCOLS; c++)
          { if (tempBoard[r][c][0] != 0)
            { redList[redTop][ROW] = r;
              redList[redTop][COL] = c;
              redList[redTop][COUNT] = tempBoard[r][c][0] / 2;
              redList[redTop++][CLOSEST] = tempBoard[r][c][1];
        } } }
  
  
        for (i = 0; i < blueTop-1; i++)
        { for (j = i; j < blueTop; j++)
          { cmp = compareRowCol(blueList[i][ROW], blueList[i][COL], blueList[j][ROW], blueList[j][COL]); 
            if (cmp == 1)
            { row = blueList[j][ROW];
              col = blueList[j][COL];
              count = blueList[j][COUNT];
              blueList[j][ROW] = blueList[i][ROW];
              blueList[j][COL] = blueList[i][COL];
              blueList[j][COUNT] = blueList[i][COUNT];
              blueList[i][ROW] = row;
              blueList[i][COL] = col;
              blueList[i][COUNT] = count;
        } } }
  
        for (i = 0; i < redTop-1; i++)
        { for (j = i; j < redTop; j++)
          { cmp = compareCloseRowCol(redList[i][CLOSEST], redList[i][ROW], redList[i][COL], redList[j][CLOSEST], redList[j][ROW], redList[j][COL]); 
            if (cmp == 1)
            { row = redList[j][ROW];
              col = redList[j][COL];
              count = redList[j][COUNT];
              close = redList[j][CLOSEST];
              redList[j][ROW] = redList[i][ROW];
              redList[j][COL] = redList[i][COL];
              redList[j][COUNT] = redList[i][COUNT];
              redList[j][CLOSEST] = redList[i][CLOSEST];
              redList[i][ROW] = row;
              redList[i][COL] = col;
              redList[i][COUNT] = count;
              redList[i][CLOSEST] = close;
        } } }

        r = 0;
        b = 0;
        found = false;
        rv.x = -1;
        maxCount = 0;
        close = MAXDIST;
        moveListTop = 0;
        while ((r < redTop) && (!found))
        { /* binsearch bluelist for redList[r][ROW]{COL] */
          lo = 0;
          hi = blueTop - 1;
          mid = 0;
          while ((lo <= hi) && (!found))
          { mid = (lo + hi) / 2;
            cmp = compareRowCol(redList[r][ROW], redList[r][COL], blueList[mid][ROW], blueList[mid][COL]);
            if (cmp > 0)
              lo = mid + 1;
            else
              if (cmp == 0)
                found = true;
              else
                hi = mid - 1;
          }
          if (found)
          { /* All those in moveList will have same Count and Closeness */
            if ( (blueList[mid][COUNT] + redList[mid][COUNT] >= maxCount)
               && (redList[r][CLOSEST] <= close))
            { if ( (blueList[mid][COUNT] + redList[mid][COUNT] > maxCount)
               || (redList[r][CLOSEST] < close))
                moveListTop = 0;
              moveList[moveListTop][ROW] = blueList[mid][ROW];
              moveList[moveListTop++][COL] = blueList[mid][COL];
              maxCount = blueList[mid][COUNT];
              close = redList[r][CLOSEST];
            }
            found = false;
          }
          r++;
        }
        if (moveListTop == 0)
        { for (b = 0; b < blueTop; b++)
          { if (blueList[b][COUNT] >= maxCount)
            { if (blueList[b][COUNT] > maxCount)
                moveListTop = 0;
              moveList[moveListTop][ROW] = blueList[b][ROW];
              moveList[moveListTop++][COL] = blueList[b][COL];
              maxCount = blueList[b][COUNT];
        } } }
        maxMoves = Math.max(moveListTop, maxMoves);
        b = (int) (Math.random() * moveListTop);
        rv.x = moveList[b][COL];
        rv.y = moveList[b][ROW];
    } }
    return(rv);
  }

  /* Returns -1 if row1,col1 < row2, col2 0 if equal +1 if greater */
  public static int compareCloseRowCol(int close1, int row1, int col1, int close2, int row2, int col2)
  { int rc1, rc2;

    rc1 = (close1 * 2000) + (row1 * 50) + col1;
    rc2 = (close2 * 2000) + (row2 * 50) + col2;
    if (rc1 == rc2)
      return(0);
    else
      if (rc1 < rc2)
        return(-1);
      else
        return(+1);

  }
  /* Returns -1 if row1,col1 < row2, col2 0 if equal +1 if greater */
  public static int compareRowCol(int row1, int col1, int row2, int col2)
  { int rc1, rc2;

    rc1 = (row1 * 100) + col1;
    rc2 = (row2 * 100) + col2;
    if (rc1 == rc2)
      return(0);
    else
      if (rc1 < rc2)
        return(-1);
      else
        return(+1);

  }

  public static int newStep()
  { int rv;
    rv = firstFreeStep;
    stepCounter++;
    maxSteps = Math.max(stepCounter, maxSteps);
    firstFreeStep = steps[firstFreeStep].nextStep;
    if (firstFreeStep == -1)
    { if (debug)
      { statusMsg.setText("Out of Memory for Steps.");
        System.out.println("StepCounter: " + stepCounter);
      }
      else
        turnMsg.setText("Out of Memory for Steps.");
    }
    steps[rv].nextStep = -1;
    return(rv);
  }

  public static void deleteStep(int stepNum)
  {
    stepCounter--;
    steps[stepNum].nextStep = firstFreeStep;
    firstFreeStep = stepNum;
  }
  
  public static void findShortestBluePaths()
  { int pathHead, pathHeadLength, pathHeadTogo, minTogo, c, row, col;
    int stepCnt, prevStep, nextStep, curStep, lastStep, targ, cp, lp;
    int row1, col1, ph, pathCnt, c1;

    minTogo = MAXROWS * MAXCOLS;

    pathCnt = 0;
    for (c = 1; c < MAXCOLS - 1; c+=2)
    { // System.out.println("Column: " + c);
      /* No need to consider row1 except for column c */
      for (c1 = 1; c1 < MAXCOLS - 1; c1+= 2)
      { if (c1 != c)
          consider[1][c1] = false;
        else
          consider[1][c1] = true;
      }
      if (board[1][c] != REDTOWER)
      { curStep = newStep();
        pathHead = curStep;
        steps[curStep].row = 1;
        steps[curStep].col = c;
        steps[curStep].stepCnt = 1;
        steps[curStep].direction = DOWN;
        steps[curStep].nextStep = -1;
        steps[curStep].prevStep = -1;
        pathHeadLength = 1;
        if (board[1][c] == NOTOWER)
          pathHeadTogo = 1;
        else
          pathHeadTogo = 0;
        visited[1][c] = true;
        optionTop = 0;
        row1 = 2;
        col1 = c;
        stepCnt = 2;
        /* check right */
        considerBlueMove(stepCnt, row1, col1+1, RIGHT);

        /* check LEFT */
        considerBlueMove(stepCnt, row1, col1-1, LEFT);

        /* check DOWN */
        considerBlueMove(stepCnt, row1+1, col1, DOWN);

        lastStep = curStep;
        stepCnt = 3;
        while (optionTop > 0)
        { optionTop--;
          /* makeMove(optionList[optionTop]); */
          while (optionList[optionTop].stepCnt <= steps[lastStep].stepCnt)
          { row = steps[lastStep].row;
            col = steps[lastStep].col;
            visited[row][col] = false;
            pathHeadLength--;
            if (board[row][col] != BLUETOWER)
              pathHeadTogo--;
            lastStep = steps[lastStep].prevStep;
            deleteStep(steps[lastStep].nextStep);
            steps[lastStep].nextStep = -1;
          }
          curStep = newStep();
          steps[lastStep].nextStep = curStep;
          steps[curStep].prevStep = lastStep;
          steps[curStep].row = optionList[optionTop].row;
          steps[curStep].col = optionList[optionTop].col;
          steps[curStep].stepCnt = optionList[optionTop].stepCnt;
          steps[curStep].direction = optionList[optionTop].direction;
          steps[curStep].nextStep = -1;
          pathHeadLength++;
          row = steps[curStep].row;
          col = steps[curStep].col;
          if (board[row][col] == NOTOWER)
            pathHeadTogo++;
          visited[row][col] = true;

          if ((steps[curStep].row == MAXROWS - 2)
           || (pathHeadTogo > minTogo)
           || (pathHeadTogo > minSteps[row][col]))
          { if ( (steps[curStep].row == MAXROWS - 2) 
              && (pathHeadTogo <= minTogo))
            { /* savePath(); */
              if (pathHeadTogo < minTogo) /* replace current shortest path(s) */
              { deleteBluePathList();
                targ = 0;
                bluePathListTop = 1;
                minTogo = pathHeadTogo;
              }
              else  /* path HeadTogo = minTogo */
              { if (bluePathListTop < MAXPATHLIST)  /* add to list of shortest path(s) */
                { targ = bluePathListTop++;
                  maxPathList = Math.max(bluePathListTop, maxPathList);
                }
                else
                  targ = MAXPATHLIST;
              }
              if (targ == MAXPATHLIST)
              { if (debug)
                  statusMsg.setText("Out of space for Blue PathList");
                else
                  turnMsg.setText("");
              }
              else
              { cp = newStep();
                bluePathList[targ].length = pathHeadLength;
                bluePathList[targ].togo = pathHeadTogo;
                bluePathList[targ].firstStep = cp;
                steps[cp].row = steps[pathHead].row;
                steps[cp].col = steps[pathHead].col;
                steps[cp].stepCnt = steps[pathHead].stepCnt;
                steps[cp].direction = steps[pathHead].direction;
                ph = steps[pathHead].nextStep;
                lp = cp;
                while (ph != -1)
                { cp = newStep();
                  steps[lp].nextStep = cp;
                  steps[cp].prevStep = lp;
                  steps[cp].row = steps[ph].row;
                  steps[cp].col = steps[ph].col;
                  steps[cp].stepCnt = steps[ph].stepCnt;
                  steps[cp].direction = steps[ph].direction;
                  ph = steps[ph].nextStep;
                  lp = cp;
                }
                steps[lp].nextStep = -1;
                /* Be sure pointers are left in right state for recursion */
            } }
/*          else   bailout */
          }
          else  /* keep going */
          { minSteps[row][col] = pathHeadTogo;
            switch (steps[curStep].direction)   /* ROW1, COL1 is location of initial BLUETOWER */
            { case UP:
                row1 = steps[curStep].row - 1;
                col1 = steps[curStep].col;
                break;
              case DOWN:
                row1 = steps[curStep].row + 1;
                col1 = steps[curStep].col;
                break;
              case LEFT:
                row1 = steps[curStep].row;
                col1 = steps[curStep].col-1;
                break;
              case RIGHT:
                row1 = steps[curStep].row;
                col1 = steps[curStep].col+1;
                break;
            }

            stepCnt = steps[curStep].stepCnt + 1;
            /* check UP */
            considerBlueMove(stepCnt, row1-1, col1, UP);

            /* check right */
            considerBlueMove(stepCnt, row1, col1+1, RIGHT);

            /* check LEFT */
            considerBlueMove(stepCnt, row1, col1-1, LEFT);

            /* check DOWN */
            considerBlueMove(stepCnt, row1+1, col1, DOWN);
          }
          lastStep = curStep;
        } /* End While (optionTop > 0) */
        while (lastStep != -1)
        { row = steps[lastStep].row;
          col = steps[lastStep].col;
          visited[row][col] = false;
          lp = steps[lastStep].prevStep;
          deleteStep(lastStep);
          lastStep = lp;
        }
      }   /* end if board[1][c] != REDTOWER */
    } /* End for c = 1 to MAXCOLS */
  }

  public static void deleteBluePathList()
  { int p, lp, cp;

    for (p = 0; p < bluePathListTop; p++)
    { cp = bluePathList[p].firstStep;
      while (cp != -1)
      { lp = cp;
        cp = steps[cp].nextStep;
        deleteStep(lp);
      }
    }
    bluePathListTop = 0;
  }

  public static void considerBlueMove(int stepCnt, int row2, int col2, int direction)
  {
    if ((consider[row2][col2])
     && (!visited[row2][col2])
     && (board[row2][col2] != REDTOWER) )
    { optionList[optionTop].stepCnt = stepCnt;
      optionList[optionTop].row = row2;
      optionList[optionTop].col = col2;
      optionList[optionTop].direction = direction;
      optionTop++;
      maxOptions = Math.max(optionTop, maxOptions);
    }
  }

  public static void findShortestRedPaths()
  { int pathHead, pathHeadLength, pathHeadTogo, minTogo, row, col;
    int stepCnt, prevStep, nextStep, curStep, lastStep, targ, cp, lp;
    int row1, col1, ph, pathCnt, r, r1;

    minTogo = MAXROWS * MAXCOLS;

    pathCnt = 0;
    for (r = 1; r < MAXROWS - 1; r+=2)
    { // System.out.println("Row: " + r);
      /* No need to consider col1 except for row r */
      for (r1 = 1; r1 < MAXROWS - 1; r1+= 2)
      { if (r1 != r)
          consider[r1][1] = false;
        else
          consider[r1][1] = true;
      }
      if (board[r][1] != BLUETOWER)
      { curStep = newStep();
        pathHead = curStep;
        steps[curStep].row = r;
        steps[curStep].col = 1;
        steps[curStep].stepCnt = 1;
        steps[curStep].direction = RIGHT;
        steps[curStep].nextStep = -1;
        steps[curStep].prevStep = -1;
        pathHeadLength = 1;
        if (board[r][1] == NOTOWER)
          pathHeadTogo = 1;
        else
          pathHeadTogo = 0;
        visited[r][1] = true;
        optionTop = 0;
        row1 = r;
        col1 = 2;
        stepCnt = 2;

        /* check UP */
        considerRedMove(stepCnt, row1-1, col1, UP);

        /* check DOWN */
        considerRedMove(stepCnt, row1+1, col1, DOWN);

        /* check right */
        considerRedMove(stepCnt, row1, col1+1, RIGHT);

        lastStep = curStep;
        stepCnt = 3;
        while (optionTop > 0)
        { optionTop--;
          /* makeMove(optionList[optionTop]); */
          while (optionList[optionTop].stepCnt <= steps[lastStep].stepCnt)
          { row = steps[lastStep].row;
            col = steps[lastStep].col;
            visited[row][col] = false;
            pathHeadLength--;
            if (board[row][col] != REDTOWER)
              pathHeadTogo--;
            lastStep = steps[lastStep].prevStep;
            deleteStep(steps[lastStep].nextStep);
            steps[lastStep].nextStep = -1;
          }
          curStep = newStep();
          steps[lastStep].nextStep = curStep;
          steps[curStep].prevStep = lastStep;
          steps[curStep].row = optionList[optionTop].row;
          steps[curStep].col = optionList[optionTop].col;
          steps[curStep].stepCnt = optionList[optionTop].stepCnt;
          steps[curStep].direction = optionList[optionTop].direction;
          steps[curStep].nextStep = -1;
          pathHeadLength++;
          row = steps[curStep].row;
          col = steps[curStep].col;
          if (board[row][col] == NOTOWER)
            pathHeadTogo++;
          visited[row][col] = true;

          if ((steps[curStep].col == MAXCOLS - 2)
           || (pathHeadTogo > minTogo)
           || (pathHeadTogo > minSteps[row][col]))
          { if ( (steps[curStep].col == MAXCOLS - 2) 
              && (pathHeadTogo <= minTogo))
            { /* savePath(); */
              if (pathHeadTogo < minTogo) /* replace current shortest path(s) */
              { deleteRedPathList();
                targ = 0;
                redPathListTop = 1;
                minTogo = pathHeadTogo;
              }
              else  /* path HeadTogo = minTogo */
              { if (redPathListTop < MAXPATHLIST)  /* add to list of shortest path(s) */
                { targ = redPathListTop++;
                  maxPathList = Math.max(redPathListTop, maxPathList);
                }
                else
                  targ = MAXPATHLIST;
              }
              if (targ == MAXPATHLIST)
              { if (debug)
                  statusMsg.setText("Out of space for Red PathList");
                else
                  turnMsg.setText("");
              }
              else
              { cp = newStep();
                redPathList[targ].length = pathHeadLength;
                redPathList[targ].togo = pathHeadTogo;
                redPathList[targ].firstStep = cp;
                steps[cp].row = steps[pathHead].row;
                steps[cp].col = steps[pathHead].col;
                steps[cp].stepCnt = steps[pathHead].stepCnt;
                steps[cp].direction = steps[pathHead].direction;
                ph = steps[pathHead].nextStep;
                lp = cp;
                while (ph != -1)
                { cp = newStep();
                  steps[lp].nextStep = cp;
                  steps[cp].prevStep = lp;
                  steps[cp].row = steps[ph].row;
                  steps[cp].col = steps[ph].col;
                  steps[cp].stepCnt = steps[ph].stepCnt;
                  steps[cp].direction = steps[ph].direction;
                  ph = steps[ph].nextStep;
                  lp = cp;
                }
                steps[lp].nextStep = -1;
                /* Be sure pointers are left in right state for recursion */
            } }
/*          else   bailout */
          }
          else  /* keep going */
          { minSteps[row][col] = pathHeadTogo;
            switch (steps[curStep].direction)   /* ROW1, COL1 is location of initial RedTOWER */
            { case UP:
                row1 = steps[curStep].row - 1;
                col1 = steps[curStep].col;
                break;
              case DOWN:
                row1 = steps[curStep].row + 1;
                col1 = steps[curStep].col;
                break;
              case LEFT:
                row1 = steps[curStep].row;
                col1 = steps[curStep].col-1;
                break;
              case RIGHT:
                row1 = steps[curStep].row;
                col1 = steps[curStep].col+1;
                break;
            }

            stepCnt = steps[curStep].stepCnt + 1;
            /* check LEFT */
            considerRedMove(stepCnt, row1, col1-1, LEFT);

            /* check UP */
            considerRedMove(stepCnt, row1-1, col1, UP);

            /* check DOWN */
            considerRedMove(stepCnt, row1+1, col1, DOWN);

            /* check right */
            considerRedMove(stepCnt, row1, col1+1, RIGHT);
          }
          lastStep = curStep;
//          if (pathCnt++ < 20)
//            displayPath(pathHead, pathHeadLength, pathHeadTogo);
        } /* End While (optionTop > 0) */
        while (lastStep != -1)
        { row = steps[lastStep].row;
          col = steps[lastStep].col;
          visited[row][col] = false;
          lp = steps[lastStep].prevStep;
          deleteStep(lastStep);
          lastStep = lp;
        }
      }   /* end if board[1][c] != BLUETOWER */
    } /* End for r = 1 to MAXROWS */
  }

  public static void deleteRedPathList()
  { int p, lp, cp;

    for (p = 0; p < redPathListTop; p++)
    { cp = redPathList[p].firstStep;
      while (cp != -1)
      { lp = cp;
        cp = steps[cp].nextStep;
        deleteStep(lp);
      }
    }
    redPathListTop = 0;
  }

  public static void considerRedMove(int stepCnt, int row2, int col2, int direction)
  {
    if ((consider[row2][col2])
     && (!visited[row2][col2])
     && (board[row2][col2] != BLUETOWER) )
    { optionList[optionTop].stepCnt = stepCnt;
      optionList[optionTop].row = row2;
      optionList[optionTop].col = col2;
      optionList[optionTop].direction = direction;
      optionTop++;
      maxOptions = Math.max(optionTop, maxOptions);
    }
  }

//  public static void main(String args[]) {
//    Bridges bridgeApp = null;
//    bridgeApp = new Bridges("Bridge Application");
//    bridgeApp.resize(MAXSCREENX, MAXSCREENY+75);
//    bridgeApp.show(true);
//  }

}
