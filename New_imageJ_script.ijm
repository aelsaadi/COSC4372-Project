//open input/image.tif 
//run script & follow prompts 
//delete low quality slices if...no clear tissue is present 
//adjust threshold based on first slide only 
//Copy & paste contents of "Results" table to excel file

//////////////////////////////////////////////////////////////////////////////////////////////////////// 
////Prepare files, settings and directory to run stereology script 
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

//Set parent output folder 
waitForUser("Open image to be analyzed. Click 'OK'. Then select the parent folder '/output_mitoc'."); 
outputDirParent = getDirectory("Choose Parent output-folder to save grids to");

//remember original file name 
original_fileID = getTitle();
//create new sub-directory to store duplicate of original files 
File.makeDirectory(outputDirParent + "/" + original_fileID);
//Set permanent output directory to newly created sub-directory 
outputDir = outputDirParent + "/" + original_fileID;
saveAs("Tiff", outputDir + "/" + original_fileID); roiManager("reset");
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////Calculate number of needed grids to sample /////////////////////////////////////////////////////////
run("Set Measurements...", "area mean area_fraction limit redirect=None decimal=2");
run("Duplicate...", "duplicate"); 
run("8-bit"); 
run("Gaussian Blur...", "sigma=20"); 
run("Measure"); 
setOption("BlackBackground", true); 
run("Convert to Mask"); 
run("Measure");
//run calculations 
Area_whole = getValue("Area"); 
Area_tissue = getValue("Area limit"); 
Areatissuepercent = getValue("%Area"); 
empty_space = 100-Areatissuepercent; 
emptyspaceforcalculationonly = empty_space/100; tissueareatosample = Area_tissue * 0.30; 
tissuegridnumber = tissueareatosample/55361.38; 
extra_grids = 2 + tissuegridnumber * emptyspaceforcalculationonly; 
number_of_grids = 1 + tissuegridnumber + extra_grids; 
total_grid_number = Area_whole/55361.38; 
selectWindow(original_fileID);
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////Create stack of random grid tiles //////////////////////////////////////////////////////////////////
mainWindow = getTitle(); 
close("\Others"); 
// close all but selected window 
h = getHeight; 
w = getWidth; 
area = w*h;
gridNum = getNumber("How many grid boxes in total:",round(total_grid_number));
boxArea = area/gridNum; // area of each grid box
boxSide = sqrt(boxArea); // side length of each grid box
numBoxY = floor(h/boxSide); // number of boxes that will fit along the width
numBoxX = floor(w/boxSide); // "" height
remainX = (w - (numBoxX*boxSide))/2; // remainder distance left when all boxes fit
remainY = (h - (numBoxY*boxSide))/2;
for (i=0; i<numBoxY; i++) { // draws rectangles in a grid, centred on the X and Y axes and adds to ROI manager
    for (j=0; j<numBoxX; j++) {
        makeRectangle((remainX + (j*boxSide)), (remainY+(i*boxSide)), boxSide, boxSide); 
        roiManager("add");
    }
}

grids = getNumber("How many grid boxes to select:", round(number_of_grids));
boxID = newArray(grids); 
gridTotal = roiManager("count"); // not neccessarily as many boxes in the grid as specified due to the constraint of # squares not fitting the width/height

for (i=0; i<grids; i++) { // generate random number to select a square in the grid, makes sure the square has not already been chosen
    randNum = round(random*gridTotal);
        for (j=0; j<boxID.length; j++) {
            if (randNum == boxID[j]) {
                randNum = round(random*gridTotal);
                j=0;
            }
        }
boxID[i] = randNum; 
selectWindow(mainWindow); 
roiManager("select", boxID[i]);
run("Duplicate...", "duplicate"); // duplicates the random square in the grid box
rename("Grid " + i);
}

for (i=0; i<grids; i++) {
    selectWindow("Grid " + i);
    pointX = round(random*boxSide); // random X coord
    pointY = round(random*boxSide); // random Y coord
    makePoint(pointX, pointY);
}

for (i=0; i<grids; i++) {
        selectWindow("Grid " + i);
        pointX = round(random*boxSide); // random X coord
        pointY = round(random*boxSide); // random Y coord
        makePoint(pointX, pointY);
        saveAs("Tiff", outputDir + "/" + original_fileID + "Grid" + i);

    }
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////Stack all images with string "Grid", save final stack, close all other images //////////////////////
run("Images to Stack", "method=[Copy (center)] name=Stack title=Grid use"); 
wait(3); 
stack = getTitle(); 
wait(3); 
selectWindow(stack); 
saveAs("Tiff", outputDir + "/" + original_fileID + "STACK"); 
run("Stack Sorter"); 
waitForUser("delete low quality slides from stack, then click 'OK'"); 
saveAs("Tiff", outputDir + "/" + original_fileID + "STACK"); 
close("\Others") roiManager("Deselect"); 
roiManager("Delete");
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////calculate area of tissue for final stack, per slice /////////////////////////////////////////////////
run("Duplicate...", "title=Duplicate_Stack duplicate"); 
Duplicate_Stack=getTitle(); 
run("Split Channels"); 
selectWindow("Duplicate_Stack (red)"); 
run("Close"); selectWindow("Duplicate_Stack (green)"); 
run("8-bit"); 
run("Gaussian Blur...", "sigma=5 stack");
stack_green = getTitle(); 
run("Threshold..."); 
setAutoThreshold("Default"); 
waitForUser("1)Click & drag bottom slider to accuratelly threshold first slice; 2)Click 'Set'; 3)Find pop-up & click 'OK'; 4)Click 'OK'");
print(original_fileID); 
selectWindow(stack_green);

//measure Area of every image in stack 
for (n=1; n<=nSlices; n++) { setSlice(n); 
print("tissue area per slice (um^2): " + getValue("Area limit")); 
}
////////////////////////////////////////////////////////////////////////////////////////////////////////

////Add data from Chunk One to Log window //////////////////////////////////////////////////////////////
print("Areawhole: " + Area_whole); 
print("Areatissue: " + Area_tissue); 
print("Percent of empty space: " + empty_space); 
print("Total area of tissue to be sampled: " + tissueareatosample); 
print("Total number of grids to sample: " + number_of_grids); 
print(original_fileID);
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////calculate area of PS6 signal per slice /////////////////////////////////////////////////////////////
selectWindow("Duplicate_Stack (blue)"); run("8-bit");
run("Gaussian Blur...", "sigma=1 stack"); 
run("Subtract Background...", "rolling=50 light stack"); 
run("Auto Threshold", "method=RenyiEntropy stack"); 
run("Analyze Particles...", "size=1.7-40 show=Masks display exclude clear include summarize add stack"); 
roiManager("Show None"); 
roiManager("Show All"); 
roiManager("Show All without labels"); 
selectWindow("Summary of Duplicate_Stack (blue)");
selectWindow("Mask of Duplicate_Stack (blue)"); 
saveAs("Tiff", outputDir + "/" + original_fileID + "STACK_BinaryMask");
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////Organize results of interest /////////////////////////////////////////////////////////////////////// 

//Copy content of the "Summary" window to the "Results" window. 
selectWindow("Summary of Duplicate_Stack (blue)"); 
text = getInfo("window.contents"); lines = split(text, "\n"); 
labels = split(lines[0], "\t"); run("Clear Results"); 
for (i=1; i  

//Copy content of the "Log" window to the "Results" window. 
selectWindow("Log"); 
text = getInfo("window.contents"); 
lines = split(text, "\n"); 
labels = split(lines[0], "\t"); for (i=1; i  
//////////////////////////////////////////////////////////////////////////////////////////////////////// 

////reset parameters and close all /////////////////////////////////////////////////////////////////////
selectWindow("Summary of Duplicate_Stack (blue)"); 
run("Close");
selectWindow("Log"); 
run("Close");
selectWindow("Results");
waitForUser("WARNING: Copy data from 'Results' table, paste to excel. Then press OK to finish"); 
waitForUser("LAST WARNING: Are you sure data from 'Results' table were copied to excel?");
selectWindow("Results"); 
run("Close");
roiManager("reset"); 
run("Clear Results"); 
selectWindow("ROI Manager"); 
run("Close");
close("*");