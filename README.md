# Architect-3D
3D Modeling software created in Visual Basic 2010
The intended purpose of this project is to design 3D architectural models.
Features included:
- Camera functions:
  -Zoom
  -360 Orbit
  -Traverse
- Drawing functions:
  -Line to
  -Square
  -Triangle
  -Curve
  -Rectangular Prism
  -Text Label(2D Text at 3D point)
  -Line Color and Fill Color
-Data storage:
  -BackBuffer  _     zone(i)(j).mesh(k)
  -FrontBuffer _       model(j).mesh(k)
  -Import and Export Model Data (modelname.dat)
  -Save and Load Zone Data (zonename.dat)
  -Save Image (image.jpg)
-Motion:
  -Frame refreshing
  -Continuous model orbit
  -Continuous model spin
  -Switch controlled Translate >X,Y,Z
  -Switch controlled Rotation  >X,Y,Z
  
Every project is a work in progress Architect 3D is a working program with few minor Glitches:
  -Rotation of models is limited to position, does not always rotate in place _
   Models rotate around a common center vector, orbits around world origin.
  -When welding models does not recalculate common origin.
  
