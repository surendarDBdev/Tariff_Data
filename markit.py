from markitdown import MarkItDown

# Filtered list based strictly on your requested DISCOMs
file = BESCOM.pdf

# Zero-config — auto-selects analyzer per file type
md = MarkItDown(cu_endpoint="<content_understanding_endpoint>")


    # Strip leading/trailing whitespaces and newlines for a cleaner terminal print
clean_name = file.strip().replace("\n", " ")
print(f"Processing: {clean_name}...")

try:
    result = md.convert(file)

    # Dynamic output naming replacing either .pdf or .docx with .md
    output_filename = file.strip().replace("\n", " ")
    if output_filename.lower().endswith(".pdf"):
        output_filename = output_filename[:-4] + ".md"
    elif output_filename.lower().endswith(".docx"):
        output_filename = output_filename[:-5] + ".md"
    else:
        output_filename = output_filename + ".md"

    # Save the markdown result to a file
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(result.markdown)

    print(f"Successfully converted and saved to {output_filename}\n")
    
except Exception as e:
    print(f"Error processing {clean_name}: {e}\n")