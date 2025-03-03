/* global PowerPoint console */

export async function insertText(text: string) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(text);
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export const insertAllSlidesAndGoToLast = async (chosenFile: string, targetSlideId: string, sourceIds: string[], formatting: boolean) => {
  await PowerPoint.run(async function (context) {
    // before
    const slidesBefore = context.presentation.slides;
    slidesBefore.load("items");
    await context.sync();

    const beforeIds = slidesBefore.items.map(slide => slide.id);

    // insertion
    context.presentation.insertSlidesFromBase64(chosenFile, {
      formatting: formatting ? "KeepSourceFormatting" : "UseDestinationTheme",
      targetSlideId: targetSlideId + "#",
      sourceSlideIds: sourceIds?.map((id) => id + "#")
    })
    await context.sync();

    // after
    const slidesAfter = context.presentation.slides;
    slidesAfter.load("items");
    await context.sync();

    const afterIds = slidesAfter.items.map(slide => slide.id);

    // find new
    const newSlides = afterIds.filter(id => !beforeIds.includes(id));
    if (newSlides.length === 0) {
      return;
    }
    const lastInsertedSlideId = newSlides[newSlides.length - 1];

    // selection
    Office.context.document.goToByIdAsync(
      lastInsertedSlideId.split("#")[0],
      Office.GoToType.Slide,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Navigation err -> ", asyncResult.error.message);
        } else {
          console.log("Success", lastInsertedSlideId);
        }
      }
    )
  })
}

function getSelectedSlideThumbnail() {
  // Office.context.document.getSelectedSlidesAsync((result) => {
  //   if (result.status === Office.AsyncResultStatus.Succeeded) {
  //     const selectedSlide = result.value.slides[0];
  //     selectedSlide.getThumbnailAsync(Office.ImageFormat.JPEG, (thumbnailResult) => {
  //       if (thumbnailResult.status === Office.AsyncResultStatus.Succeeded) {
  //         const thumbnailUrl = thumbnailResult.value;
  //         console.log("Thumbnail URL: ", thumbnailUrl);
  //         // You can now use the thumbnail URL as needed        } else {
  //         console.error("Failed to get thumbnail: ", thumbnailResult.error.message);
  //       }
  //     });
  //   } else {
  //     console.error("Failed to get selected slide: ", result.error.message);
  //   }
  // })
};