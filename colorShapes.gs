function onInstall() {
  onOpen();
}

function onOpen() {
  SlidesApp.getUi()
    .createAddonMenu()
    .addSubMenu(
      SlidesApp.getUi()
        .createMenu("Color Menu")
        .addItem("Convert To A1", "convertToA1")
        .addItem("Convert To A2", "convertToA2")
        .addItem("Convert To A3", "convertToA3")
        .addItem("Convert To A4", "convertToA4")
        .addItem("Convert To A5", "convertToA5")
        .addItem("Convert To A6", "convertToA6")
        .addItem("Convert To A7", "convertToA8")
        .addItem("Convert To A8", "convertToA9")
        .addItem("Convert To A9", "convertToA9")
    )
    .addToUi();
}

//set your colors here
var colora1 = "##ff0000";
var colora2 = "";
var colora3 = "";
var colora4 = "";
var colora5 = "";
var colora6 = "";
var colora7 = "";
var colora8 = "";
var colora9 = "";

function convertToA1() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora1);
    }
  }
}

function convertToA2() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora2);
    }
  }
}

function convertToA3() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora3);
    }
  }
}

function convertToA4() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora4);
    }
  }
}

function convertToA5() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora5);
    }
  }
}

function convertToA6() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora6);
    }
  }
}

function convertToA7() {
  const slides = SlidesApp.getActivePresentation();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora7);
    }
  }
}

function convertToA8() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora8);
    }
  }
}

function convertToA9() {
  const slides = SlidesApp.getActivePresentation().getSlides();
  var shapes = slides.map((slide) => slide.getShapes());

  if (shapes.length > 0) {
    shapes = shapes.flat(1);
    for (let shape of shapes) {
      shape.getFill().setSolidFill(colora8);
    }
  }
}
