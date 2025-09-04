# PPTX Translation Pipeline (JA→EN)

This repository contains a GitHub Action to automatically translate Japanese PowerPoint files (.pptx) into English. It preserves the original layout, generates several useful artifacts for quality assurance, and uploads the results to a specified Google Drive folder.

## How to Use

This workflow is triggered manually via the `workflow_dispatch` event. You can run it from the Actions tab of the GitHub repository.

### Inputs

The action can be triggered with one of two input sources:

1.  **From Google Drive (recommended for large files):**
    *   `drive_input_folder_id`: The ID of the Google Drive folder containing the `.pptx` file you want to translate. The action will automatically pick the most recently modified `.pptx` file in that folder.
    *   `file_name_regex`: (Optional) A regular expression to filter the files in the Drive folder. Defaults to `.*\.pptx$`, which matches any file ending in `.pptx`.

2.  **From the Repository:**
    *   `repo_file_path`: The path to a `.pptx` file located within this repository (e.g., `decks/source/my_deck.pptx`).

Leave the unused input source blank.

### Outputs

Upon successful completion, the workflow produces the following outputs:

1.  **GitHub Artifacts:** A `.zip` file named `translated-pptx-and-artifacts` containing:
    *   `output_en.pptx`: The translated English version of the presentation.
    *   `bilingual.csv`: A CSV file mapping each Japanese string to its English translation for easy review.
    *   `translation_cache.json`: A cache of all translations. This file is used to avoid re-translating the same text, saving time and cost.
    *   `audit.json`: A JSON report containing statistics like the number of Japanese characters before and after translation.
    *   A copy of the `glossary.json` and the Python scripts used in the run.

2.  **Google Drive:** All the generated artifacts are also uploaded to a folder named `translation` in the root of your Google Drive.

## Operational Notes

### Glossary Updates

To improve translation consistency for specific terms, you can add entries to the `glossary.json` file. 

*   **To add or change a term:** Edit `glossary.json`, commit the change to the repository, and re-run the workflow. 
*   **Cached Translations:** The system uses the `translation_cache.json` to avoid re-translating text. If you update a glossary term, only new or changed strings will be re-translated. To force a full re-translation, you would need to manually delete the `translation_cache.json` file before running the action (though this is not typically necessary).

### Known Limitations

*   **Embedded Text:** Text embedded within images, charts, or other non-text objects cannot be translated by this script.
*   **Text Overflow:** English text is often longer than the original Japanese. While the script preserves the layout, some text boxes may overflow. You can manually fix this in the output `.pptx` file by enabling the "Shrink text on overflow" option or by adjusting font sizes.

## QA and Acceptance Checklist

- [ ] **Action Completes:** The GitHub Action workflow finishes without any errors.
- [ ] **Artifacts Present:** The `translated-pptx-and-artifacts` zip file is available for download and contains all the expected files.
- [ ] **Drive Upload:** The output files are present in the `translation` folder in Google Drive.
- [ ] **Audit Report:** `audit.json` shows `jp_chars_after` is close to zero. Any remaining characters are likely from untranslatable embedded text.
- [ ] **Spot-Check:** Review headings and bullet points in `output_en.pptx` for correct tone and accuracy.
- [ ] **Glossary Applied:** Verify that the terms from `glossary.json` have been consistently applied in the translated presentation.

## Common Pitfalls and Fixes

*   **Drive Upload Fails:** This usually means the Google Service Account does not have permission to write to the target folder. Ensure you have shared the `translation` folder (or your Drive root) with the service account's email address.
*   **No `.pptx` Found in Drive:** The `drive_input_folder_id` may be incorrect, or the `file_name_regex` might be too strict. Double-check the folder ID and the regex.
*   **JSON Parse Error from Model:** This is rare but can happen. Re-running the job will often fix it. If it persists, you can try lowering the `batch` size input.

## Maintenance

*   **Model:** The OpenAI model can be updated via the `model` input in the workflow dispatch menu (defaults to `gpt-4o`).
*   **Dependencies:** The Python dependencies are pinned in the `.github/workflows/translate-pptx.yml` file and can be updated as needed.

## Optional Enhancements (Future Considerations)

*   **Notifications:** Add Slack or email notifications on workflow completion, including direct links to the uploaded Drive files.
*   **Scheduled Runs:** Configure the workflow to run on a schedule (e.g., nightly) to process new decks in an "inbox" folder on Google Drive.
*   **OCR for Images:** Integrate an OCR tool like Tesseract to detect and translate text embedded in images.
