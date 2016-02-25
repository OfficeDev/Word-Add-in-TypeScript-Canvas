/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

class App {

    constructor() { 
        // The initialize function is run each time the page is loaded.
        Office.initialize = (reason) => {
            $(document).ready(() => {
                // Use this to check whether the new API is supported in the Word client.
                if (Office.context.requirements.isSetSupported("WordApi", 1.2)) {

                    console.log('This code is using Word 2016 or greater.');
                    
                    // Setup the event handlers for UI.
                    $('#loadSelectedImage').click(this.loadSelectedImageHandler.bind(this));
                    $('#insertImageAtSelection').click(this.insertImageHandler.bind(this));

                    // Scale the size of the canvas so that it scales  
                    // when a user resizes the add-in.
                    window.addEventListener('resize', this.resizeCanvas.bind(this), false);
            
                    // Setup the canvas event listener(s).
                    this.initCanvas();
                    
                    // Automatically load an image if one is selected in Word when the task pane is launched.
                    this.loadSelectedImageHandler();

                } else {
                    // Just letting you know that this code will not work with your version of Word.
                    console.log('This add-in requires the WordAPI 1.2 requirement set or greater. Check your version of Word and the requirement set version.');
                }
            });
        };
    }

    /*********************/
    /* Globals           */
    /*********************/

    private _calloutEnabled: boolean = false; // we only want to add callout when an image is loaded.
    private _calloutNumber: number; // set/reset when an image is loaded.
    private _resizeRatio: number; // set when an image has been loaded in to the canvas.
    private _windowWidth: number; // we are setting the canvas width to the window width.
    private _image; // the image added to the canvas.
 
    /*********************/
    /* Canvas functions */
    /*********************/

     
    /**
    * Initialize the canvas with the click event. Click event inserts callouts into the canvas image. 
    */
    initCanvas(): void {

        this._windowWidth = window.innerWidth;

        var canvas = document.getElementById('canvas') as HTMLCanvasElement;

        var ctx = canvas.getContext("2d") as CanvasRenderingContext2D;

        // Add callouts when the user clicks in the canvas.
        ctx.canvas.addEventListener('click', (event) => {

            // Let's make sure that we have an image loaded before
            // we add callouts to the canvas.
            if (this._calloutEnabled) {

                // Increment callout number. We will use this later when we stub out
                // descriptions for the callouts by using the Word JS API.
                this._calloutNumber++;

                // Get the bounds of the canvas element in relationship to the top-left of the viewport.
                // We will get the coordinates in canvas, not the window.
                var canvasBounds = canvas.getBoundingClientRect();

                // Use the event coordinates, canvas boundaries, and the window width
                // to get the coordinates where the callouts can be placed.
                var height = this._windowWidth * this._resizeRatio;
                var mouseX = (event.clientX - canvasBounds.left) * canvas.width / this._windowWidth;
                var mouseY = (event.clientY - canvasBounds.top) * canvas.height / height;
            
                // Draw circle for the callout.
                var radius = 12;
                ctx.fillStyle = 'red';
                ctx.beginPath();
                ctx.arc(mouseX, mouseY, radius, 0, Math.PI * 2, true);
                ctx.closePath();
                ctx.fill();

                // Insert the callout number in the circle.
                var width = ctx.measureText(this._calloutNumber.toString());
                ctx.font = 'bold 16px calabri ';
                ctx.fillStyle = 'white';
                ctx.textAlign = 'center';
                ctx.fillText(this._calloutNumber.toString(), mouseX, mouseY + (radius / 3)); // this last argument is approximately correct for placement.
            }
        });
    }

    /**
    * Changes the canvas size according to the window width and the image aspect ratio. 
    */
    resizeCanvas(): void {
    
        // Canvas must fit width of add-in.
        this._windowWidth = window.innerWidth;

        // Set the resize ratio only if it hasn't already been captured, 
        // and only if there is an image loaded in the add-in.
        if (!this._resizeRatio && this._image)
            this._resizeRatio = this._image.height / this._image.width;

        // Resize the canvas only if there is an image loaded into it.
        if (this._image) {
            var height = this._windowWidth * this._resizeRatio;
            var canvas = document.getElementById('canvas');
            canvas.style.width = this._windowWidth + 'px';
            canvas.style.height = height + 'px';
        }
    }


    /**
     * Loads the image into the HTML canvas. loadSelectedImageHandler() checks whether you have an image.
     * @param base64EncodedImage The image to load into the canvas.
     */
    loadImageIntoCanvas(base64EncodedImage): void {

        // Callouts should only be added once the image is loaded into canvas.
        this._calloutEnabled = false;
        
        // Create an image and load it onto the canvas, set the canvas to the image
        // dimensions, and draw it on the canvas.
        this._image = new Image();
        this._image.onload = () => {

            var canvas = document.getElementById('canvas') as HTMLCanvasElement;
            var ctx = canvas.getContext("2d");

            canvas.height = this._image.height;
            canvas.width = this._image.width;
            ctx.drawImage(this._image, 0, 0);
         
            // Reset the this._calloutNumber when I load a new image.
            this._calloutNumber = 0;

            // Enable adding callouts to the canvas.
            this._calloutEnabled = true;
            
            // Make the canvas scale to the window.
            this.resizeCanvas();
        };

        // ASSUMPTION: we are assuming only png files. You will need to determine file type.
        // Load the image we got from Word.
        this._image.src = "data:image/png;base64," + base64EncodedImage.value;
    }

    /*********************/
    /* Word JS functions */
    /*********************/

    /**
    * Load the the selected image from Word into the add-in. This assumes that a single image was selected. 
    */
    loadSelectedImageHandler(): void {
        Word.run((context) => {

            // Create a proxy object for the range that is assumed to contain an image.
            var imageRange = context.document.getSelection();

            // Load the selected range.
            context.load(imageRange, 'inlinePictures');

            // Synchronize the document state by executing the queued commands, 
            // and return a promise to indicate task completion.
            return context.sync()
                .then(() => {
                    // If there is more than one inline picture, then we need to tell the user to choose a single picture.
                    if (imageRange.inlinePictures.items.length === 1) {

                        // Queue a command to get the image source. 
                        var imageString = imageRange.inlinePictures.items[0].getBase64ImageSrc();

                        // Synchronize the document state by executing the queued commands, 
                        // and return a promise to indicate task completion.
                        return context.sync().then(() => {
                            this.loadImageIntoCanvas(imageString);
                        });

                    }
                    // 
                    else {
                        throw "You need to select a single image."
                    }

                });
        })
            .catch((error) => {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }
  
    /**
    * Insert the contents of the canvas into the Word document. 
    */
    insertImageHandler(): void {

        // Only insert the contents of the canvas if we an image in it.
        if (this._image) {

            var canvas = document.getElementById('canvas') as HTMLCanvasElement;

            // Get the data URL for the image in the canvas. 
            var pngDataUrl = canvas.toDataURL(); // data uri scheme

            // Extract the encoding format information. Word only accepts the base64 content. 
            // ASSUMPTION: that this is a png file.
            var base64ImgString = pngDataUrl.replace('data:image/png;base64,', '');

            Word.run((context) => {

                // Create a proxy object for the range at the current selection.
                var imageRange = context.document.getSelection();

                // Load the selected range.
                context.load(imageRange, 'text');

                // Synchronize the document state by executing the queued commands, 
                // and return a promise to indicate task completion.
                return context.sync()
                    .then(() => {

                        // Queue a command to insert the image into the document.
                        var insertedImage = imageRange.insertInlinePictureFromBase64(base64ImgString, Word.InsertLocation.replace);

                        // Queue a command to navigate the UI to the insert picture.
                        insertedImage.select();

                        // Queue an indefinite number of commands to insert paragraphs 
                        // based on the number of callouts added to the image. 
                        if (this._calloutNumber > 0) {
                            var lastParagraph = insertedImage.insertParagraph('Here are your callout descriptions:', Word.InsertLocation.after) as Word.Paragraph;

                            for (var i = 0; i < this._calloutNumber; i++) {
                                lastParagraph = lastParagraph.insertParagraph((i + 1) + ') [enter callout description].', Word.InsertLocation.after);
                            }
                        }
                    })
                    // Synchronize the document state by executing the queued commands.
                    .then(context.sync);
            })
                .catch((error) => {
                    console.log('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });
        }
    }
}

var app = new App();