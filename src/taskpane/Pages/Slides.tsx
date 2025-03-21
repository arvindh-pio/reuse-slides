import React from "react";
import { insertAllSlidesAndGoToLast } from "../taskpane";
import { makeStyles } from "@fluentui/react-components";
import { ISlide } from "../components/App";

interface IPreview {
  previews: string[];
  base64: string;
  sourceSlideIds: ISlide[];
  formatting: boolean;
  setFormatting: React.Dispatch<React.SetStateAction<boolean>>;
  handleBack: () => void;
}

const useStyles = makeStyles({
  previewContainer: {
    display: "flex",
    flexDirection: "column",
    margin: "0 auto",
    alignItems: "center",
    gap: "10px"
  },
  slide: {
    border: "1px solid black",
    width: "200px",
    height: "100px",
    borderRadius: "6px",
    display: "flex",
    justifyContent: "center",
    alignItems: "center"
  },
  insertBtn: {
    background: "none",
    cursor: "pointer",
    border: "0",
    padding: "8px 16px",
    borderRadius: "8px",
    transition: "all 0.3s",
    fontFamily: "system-ui",
    ":hover": {
      backgroundColor: "lightgray"
    }
  },
  formatDiv: {
    display: "flex",
    gap: "4px"
  }
})

const Slides = (props: IPreview) => {
  const styles = useStyles();
  const { previews, base64, sourceSlideIds, formatting, setFormatting, handleBack } = props;

  // handlers
  const handleInsert = async (slideId: string, base64: string, insertAll: boolean = false) => {
    const targetSlideId = await getSelectedSlideId();
    const sourceIds = insertAll ? sourceSlideIds?.map((slide) => slide?.slideId) : [slideId];
    await insertAllSlidesAndGoToLast(base64, targetSlideId, sourceIds, formatting);
  }

  const getSelectedSlideId = () => {
    return new OfficeExtension.Promise<string>(function (resolve, reject) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        try {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(console.error(asyncResult.error.message));
          } else {
            resolve((asyncResult.value as any).slides[0].id);
          }
        } catch (error) {
          reject(console.log(error));
        }
      })
    })
  }

  return (
    <div className={styles.previewContainer}>
      <div style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        width: "100%",
        margin: "1rem 0 0 0"
      }}>
        <h3 style={{ margin: "0" }}>Presentation</h3>
        <p style={{ textDecoration: "underline", margin: "0", cursor: "pointer" }} onClick={handleBack}>Back</p>
      </div>
      <div style={{
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        width: "100%"
      }}>
        {previews?.length > 0 || sourceSlideIds?.length > 0 &&
          <button className={styles.insertBtn} onClick={() => handleInsert("", base64, true)}>Insert all</button>}
        {/* source formatting */}
        <div className={styles.formatDiv}>
          <input
            type="checkbox"
            name="formatting"
            id="formatting"
            checked={formatting}
            disabled={!sourceSlideIds?.length}
            onChange={() => setFormatting((prev) => !prev)} />
          <label htmlFor="formatting">Keep source formatting</label>
        </div>
      </div>
      {previews?.length > 0 ? previews?.map((preview, idx) => {
        return (
          <img
            key={preview + idx}
            src={preview}
            alt="Slide image"
            onClick={() => handleInsert(sourceSlideIds?.[idx]?.slideId, base64)} />
        )
      }) : (
        sourceSlideIds?.map((slide, idx) => {
          return (
            <div
              className={styles.slide}
              key={slide?.index + "_" + idx}
              onClick={() => handleInsert(sourceSlideIds?.[idx]?.slideId, base64)}
            >Slide - {slide?.index}</div>
          )
        })
      )}
    </div>
  );
};

export default Slides;
