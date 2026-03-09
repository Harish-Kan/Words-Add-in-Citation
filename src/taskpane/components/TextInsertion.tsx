  import * as React from "react";
import { useState } from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";

/* global Word */



const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },

});

const TextInsertion: React.FC = () => {

  const styles = useStyles();

const [citation, setCitation] = useState<any>(null);

const handleAnalyze = async () => {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");

    await context.sync();

    const selectedText = range.text;

    if (!selectedText || selectedText.trim().length === 0) {
      setCitation({ error: "Please select text before analyzing." });
      return;
    }

    console.log("Selected text:", selectedText);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          text: selectedText,
          document_id: "trustops-handbook-v1",
          user_id: "candidate_1",
        }),
      });

      const data = await response.json();

      range.font.highlightColor = "#FFFF00";
      await context.sync();

      console.log("Citation response:", data);

      // Check for null or low-confidence response
      if (!data || data.confidence < 0.2) {
        setCitation({ error: "Unable to generate citation. Try refining selection." });
      } else {
        setCitation(data);
      }
    } catch (error) {
      console.error("API request failed:", error);
      setCitation({ error: "API request failed. Please try again." });
    }
  });
};

  return (
    <div className={styles.textPromptAndInsertion}>

      <Field className={styles.instructions}>Highlight the text and Click to anaylze .</Field>
      <Button appearance="primary" size="large" onClick={handleAnalyze}>
        Analyze Selection
      </Button>

        {citation && (
          <div style={{ marginTop: "20px", textAlign: "center" }}>
            {citation.error ? (
              <p>{citation.error}</p>
            ) : (
              <>
                <h3>Citation</h3>
                <p>{citation.citation_text}</p>
                <p>Confidence: {citation.confidence}</p>
                <a href={citation.url} target="_blank">
                  View Source
                </a>
              </>
            )}
          </div>
        )}

    </div>
  );
};

export default TextInsertion;
