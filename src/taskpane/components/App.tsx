import React, { useState } from "react";
import { makeStyles } from "@fluentui/react-components";
import FileInput from "./FileInput";
import Previews from "./Previews";

export interface ISlide {
  index: number;
  slideId: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "1rem 2rem"
  },
});

const App: React.FC = () => {
  // styles
  const styles = useStyles();

  // states
  const [file, setFile] = useState<File | null>(null);
  const [previews, setPreviews] = useState<string[]>([]);
  const [sourceSlideIds, setSourceSlideIds] = useState<ISlide[]>([]);
  const [base64, setBase64] = useState<string | null>(null);
  const [formatting, setFormatting] = useState(false);

  return (
    <div className={styles.root}>
      <FileInput
        setFile={setFile}
        setPreviews={setPreviews}
        setSourceSlideIds={setSourceSlideIds}
        setBase64={setBase64}
        setFormatting={setFormatting}
        sourceSlidesLength={sourceSlideIds?.length}
        formatting={formatting} />
      <Previews base64={base64} previews={previews} sourceSlideIds={sourceSlideIds} formatting={formatting} />
    </div>
  );
};

export default App;
