### Document Similarity Pipeline
A hybrid NLP + rule-based system for comparing structured documents section-wise to verify content consistency, even when section names or writing styles differ.

This project enables users to verify if sections of a primary document are present, accurate, and correctly reflected across one or more target documents.

### Features
Section-wise Comparison: Match, Partial Match, or No Match outcomes for each section.

Handling Variations: Supports different section names and condensed/summarized sections.

Semantic Matching: Uses Sentence Transformers for deep meaning-based comparisons.

Fast Search: Utilizes FAISS for rapid similarity retrieval over large documents.

Bullet Points and Tables: Bullet points and embedded tables are also compared when needed.

Rule-Based Support: Custom rules for pre-defined section mappings between documents.

### Technology Stack
Sentence Transformers (paraphrase-MiniLM-L6-v2) for text embeddings.

FAISS for similarity search indexing.

Python libraries: docx, numpy, regex, torch.

## How It Works
Document Parsing:
Extract headings, subheadings, bullet points, and tables from Word documents.

Embedding Generation:
Encode sections into dense semantic vectors using Sentence Transformers.

Similarity Search:
Index target document embeddings using FAISS and retrieve closest matches for each reference section.

Rule-Based Mapping:
Apply manual mappings between known related sections across documents when names differ.

Result Interpretation:
Determine the match type based on similarity score thresholds:

Full Match (score ≥ 0.8)

Partial Match (0.5 ≤ score < 0.8)

No Match (score < 0.5)

Install dependencies:

pip install sentence-transformers faiss-cpu python-docx numpy
Prepare your documents:
Save your reference and target .docx files.

Review Results:
View match type and similarity scores for each section!
