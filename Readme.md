# SplineTravel
A g-code filter to smooth out 3D-printing process

## What does it do?
It processes G-code programs for 3D printing and replaces straight-line travel moves with smooth curved moves, aimed to avoid the extruder coming to a complete stop before and after the move.

Also, it features a seam prevention technique, similar to Slic3r's "Wipe while retracting", but more comprehensive

## How to use
Slice your model with a slicer of your choice. Write g-code to a file. Feed this file to SplineTravel. SplineTravel will write its output g-code into another file. Use this file for printing. See more in Wiki.

## Installing
There is a pre-built executable right in the repository. It is more-or-less stand-alone (provided you have VB6 runtime libraries, which are included in most versions of Windows). There is no installer, so far.

## Building or running from source code
SplineTravel is written in Visual Basic 6. To use the source code, all that is needed is a working installation of VB6 IDE. SplineTravel has no dependencies on external libraries other than VB6 runtime.


## Links
project page on hackaday: <https://hackaday.io/project/7045-splinetravel>    
Slic3r, the slicer this project is mostly tested with: <http://slic3r.org/>
