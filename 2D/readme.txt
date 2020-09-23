Matrix Lessons using Asteroids
==============================

Comments
=========
This project originally started off as a lesson in Matrix Multiplication using 3D vectors; And by "3D vectors", I don't mean 
3D computer graphics... I mean "vectors".  I was actually half way through coding a 3D computer graphics program, when I 
decided I needed a refresher course in matrix multiplication... for you see to do 3D graphics, I actually use "4D vectors". 
Confused?  Well... so was I, so I decided to drop back a dimension and learn all over from "2D graphics using 3D vectors", so 
that I could understand "3D graphics using 4D vectors".

Anyway... so the matrix math is going well, when I think to myself, "Hmmm... maybe I should create an Asteroids game"
Well... I'm half-way through debugging the AI routine for the enemy spaceships, when I decide that with all of my debugging 
graphics turned on, it looked kind of cool, so I thought you would want to see it before I pull out all the debugging 
graphics. The actually Asteroids game will probably be finished in another month or so. I don't want to release it until I've 
got a kick-ass A. I routine for the enemy space ships (basically I want to make computer controlled ships, just as smart as 
human players)


Game Play
=========
You will have to code this yourself, or wait for the finished product.


Keyboard
========
* The space bar changes the levels.
* You control only one of the red thing’s with the cursor keys.
* The P button Pauses the simulation (the timer control keeps going... I will put this to good use later)
* Double-click anywhere on the form to reset it's aspect ratio 1:1
* Mouse-Down on the form to freeze the Timer control (similar to a Pause button)


Compiler
========
* Change the "Conditional Compiler Options" to turn off/on the debuggin vectors.
  (Must be changed in the Project Properties area)
	gcShowVectors = -1  or
	gcShowVectors = 0

* Compile for speed (of course), and don't forget to play around with the Timer control's interval.


Features to look out for...
===========================
* There are not too many comments at this stage, however when I finish the game, then I will fully document it.
  If you have any questions, just ask and I will explain.

* Lots of cool stuff, scattered throughout the code!
  I personally love the routine I worte to create a random asteroid.... this is really what started the whole project.

* Matrix Concatenation using Matrix Multiplication
  (Note: The order in which the Matrices are multiplied together.)

	matResult = MatrixIdentity
	matResult = MatrixMultiply(matResult, m_matScale)
	matResult = MatrixMultiply(matResult, matRotationAboutZ)
	matResult = MatrixMultiply(matResult, matTranslate)
	matResult = MatrixMultiply(matResult, m_matViewMapping)

* Matrix * Vector multiplication
  (This is the fun part, that changes our 3D vector into 2D screen space)

	For intJ = LBound(.Vertex) To UBound(.Vertex)
		.TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
	Next intJ


Feedback
========
If you've got any questions, praise or comments then send my an e-mail.
If you feel so inclined to vote for this code on Planet-Source-Code, then that would be good too although not necessary.


Peter Wilson
peter@midar.com
http://www.midar.com/vblessons/default.asp

