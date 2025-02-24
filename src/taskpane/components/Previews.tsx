import React, { useState } from "react";
import { insertAllSlidesAndGoToLast } from "../taskpane";
import { ISlide } from "./App";
import { makeStyles } from "@fluentui/react-components";

interface IPreview {
    previews: string[];
    base64: string;
    sourceSlideIds: ISlide[];
    formatting: boolean;
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
        border: "1px solid crimson",
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
        ":hover": {
            backgroundColor: "lightgray"
        }
    }
})

const Previews = (props: IPreview) => {
    const styles = useStyles();
    const { previews, base64, sourceSlideIds, formatting } = props;
    // const [selectedSlides, setSelectedSlides] = useState();

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
            {previews?.length > 0 || sourceSlideIds?.length > 0 && 
                <button className={styles.insertBtn} onClick={() => handleInsert("", base64, true)}>Insert all</button>}
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
                        >{slide?.index}</div>
                    )
                })
            )}
        </div>
    );
};

export default Previews;
