"""
Execution Error Taxonomy Classifier
===================================

Reference-based fuzzy classification of Python execution errors.

Method:
    1. Extract the meaningful error text from `exec_feedback`.
    2. Normalize dynamic values such as numbers and variable names.
    3. Compare the normalized error against manually audited references.
    4. Calculate a weighted RapidFuzz similarity score.
    5. Assign the taxonomy of the highest-scoring reference when
       similarity is strictly greater than 80%.
    6. Otherwise assign UNCLASSIFIED.

Taxonomy:
    - Interface & Formatting Errors
    - Data Type & Casting Errors
    - Initialization & Scope Errors
    - Math & Transformation Errors
    - UNCLASSIFIED

The script is designed to run as a normal Python program, not only in
Google Colab.

Usage:
    python execution_error_classifier.py

Input:
    Place an Excel file containing an `exec_feedback` column in the
    same directory and set INPUT_FILE below.

Output:
    Execution_Error_Taxonomy_Fuzzy_Analysis.xlsx
"""

from pathlib import Path
import ast
import re

import pandas as pd
from rapidfuzz import fuzz


# =============================================================================
# 1. CONFIGURATION
# =============================================================================

# Change this to the name/path of your input Excel file.
INPUT_FILE = "input_dataset.xlsx"

# Name of the generated Excel workbook.
OUTPUT_FILE = "Execution_Error_Taxonomy_Fuzzy_Analysis.xlsx"

# A classification is accepted only when the best similarity is
# strictly greater than this value.
FUZZY_THRESHOLD = 80.0


# =============================================================================
# 2. MANUALLY AUDITED REFERENCE BANK
# =============================================================================
#
# These examples are manually audited reference patterns.
# They are NOT automatically generated rules.
#
# New audited examples can be added here as the taxonomy evolves.
# =============================================================================

REFERENCE_EXAMPLES = [

    # -------------------------------------------------------------------------
    # INTERFACE & FORMATTING ERRORS
    # -------------------------------------------------------------------------

    {
        "reference_id": "IF01",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Test Case {'S1': [60, 70, 80], 'S2': [90, 95, 100]}
        Expected Output [75.0, 82.5, 90.0]
        Actual Output : 70.0
        """,
    },

    {
        "reference_id": "IF02",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Runtime issue: TypeError:
        main() takes 1 positional argument but 3 were given
        """,
    },

    {
        "reference_id": "IF03",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Runtime issue: TypeError:
        main() takes 0 positional arguments but 2 were given
        """,
    },

    {
        "reference_id": "IF04",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Runtime issue: TypeError:
        main() missing 1 required positional argument: 'y1'
        """,
    },

    {
        "reference_id": "IF05",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Runtime issue: TypeError:
        main() missing 1 required positional argument: 'odd'
        """,
    },

    {
        "reference_id": "IF06",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Runtime issue: TypeError:
        main() missing 5 required positional arguments:
        'y1', 'x2', 'y2', 'x3', and 'y3'
        """,
    },

    {
        "reference_id": "IF07",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Test Case : (0, 0, 5, 5)
        Expected Output : (1.0, 0.0)
        Actual Output : ('1.000000', '0.000000')
        """,
    },

    {
        "reference_id": "IF08",
        "taxonomy": "Interface & Formatting Errors",
        "example": """
        Test Case(0, 0, 1, 1)
        Expected Output(1.0, 0.0)
        Actual Output : {'slope': 1.0, 'intercept': 0.0}
        """,
    },

    # -------------------------------------------------------------------------
    # DATA TYPE & CASTING ERRORS
    # -------------------------------------------------------------------------

    {
        "reference_id": "DT01",
        "taxonomy": "Data Type & Casting Errors",
        "example": """
        Runtime issue: TypeError: 'int' object is not iterable
        """,
    },

    {
        "reference_id": "DT02",
        "taxonomy": "Data Type & Casting Errors",
        "example": """
        Test Case(-1, -1, 2, 1)
        Expected Output(0.666667, -0.333333)
        Actual Output : (0.666667, -0.33333333333333337)
        """,
    },

    # -------------------------------------------------------------------------
    # INITIALIZATION & SCOPE ERRORS
    # -------------------------------------------------------------------------

    {
        "reference_id": "IS01",
        "taxonomy": "Initialization & Scope Errors",
        "example": """
        Runtime issue: UnboundLocalError:
        cannot access local variable 'larger_number'
        where it is not associated with a value
        """,
    },

    {
        "reference_id": "IS02",
        "taxonomy": "Initialization & Scope Errors",
        "example": """
        Runtime issue: UnboundLocalError:
        cannot access local variable 'char'
        where it is not associated with a value
        """,
    },

    {
        "reference_id": "IS03",
        "taxonomy": "Initialization & Scope Errors",
        "example": """
        Runtime issue: UnboundLocalError:
        cannot access local variable 'sentence'
        where it is not associated with a value
        """,
    },

    {
        "reference_id": "IS04",
        "taxonomy": "Initialization & Scope Errors",
        "example": """
        Runtime issue: NameError:
        name 'a' is not defined
        """,
    },

    # -------------------------------------------------------------------------
    # MATH & TRANSFORMATION ERRORS
    # -------------------------------------------------------------------------

    {
        "reference_id": "MT01",
        "taxonomy": "Math & Transformation Errors",
        "example": """
        Runtime issue: ZeroDivisionError:
        division by zero
        """,
    },

    {
        "reference_id": "MT02",
        "taxonomy": "Math & Transformation Errors",
        "example": """
        Runtime issue: ZeroDivisionError:
        float division by zero
        """,
    },
]


# =============================================================================
# 3. TEXT NORMALIZATION
# =============================================================================

def normalize_text(text):
    """
    Perform basic text cleaning before error extraction/matching.

    Operations:
        - Handle missing values.
        - Convert HTML line breaks to spaces.
        - Collapse repeated whitespace.
        - Convert text to lowercase.

    Parameters
    ----------
    text : object
        Raw text value.

    Returns
    -------
    str
        Cleaned text.
    """
    if text is None:
        return ""

    if pd.isna(text):
        return ""

    text = str(text)

    # Convert HTML line breaks to normal spaces.
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)

    # Remove excessive whitespace.
    text = re.sub(r"\s+", " ", text)

    return text.lower().strip()


# =============================================================================
# 4. ERROR EXTRACTION
# =============================================================================

def extract_error_text(value):
    """
    Extract the meaningful error portion from raw exec_feedback.

    Many records have a structure similar to:
        {'test_case': (...), 'error': 'Runtime issue: TypeError: ...'}

    The function first attempts safe Python-literal parsing. If that fails,
    it uses regex-based fallbacks.

    Parameters
    ----------
    value : object
        Raw exec_feedback value.

    Returns
    -------
    str
        Extracted and normalized error text.
    """
    if value is None or pd.isna(value):
        return ""

    text = str(value).strip()

    if not text:
        return ""

    # -------------------------------------------------------------------------
    # STEP 1: Try to parse dictionary-like execution feedback safely.
    # ast.literal_eval() is preferred over eval() because it does not execute
    # arbitrary Python code.
    # -------------------------------------------------------------------------
    try:
        parsed = ast.literal_eval(text)

        if isinstance(parsed, dict) and "error" in parsed:
            return normalize_text(str(parsed["error"]))

    except (ValueError, SyntaxError, TypeError):
        pass

    # -------------------------------------------------------------------------
    # STEP 2: Fallback for dictionary-like text when literal parsing fails.
    # -------------------------------------------------------------------------
    error_marker = re.search(
        r"""['"]error['"]\s*:\s*""",
        text,
        flags=re.IGNORECASE,
    )

    if error_marker:
        error_text = text[error_marker.end():].strip()

        # Remove outer braces if present.
        error_text = error_text.strip("{} ")

        # Remove surrounding quote only if it encloses the complete value.
        if (
            len(error_text) >= 2
            and error_text[0] in ("'", '"')
            and error_text[-1] == error_text[0]
        ):
            error_text = error_text[1:-1]

        return normalize_text(error_text)

    # -------------------------------------------------------------------------
    # STEP 3: If the feedback contains "Runtime issue", keep that section.
    # -------------------------------------------------------------------------
    runtime_match = re.search(
        r"runtime issue\s*:?.*",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )

    if runtime_match:
        return normalize_text(runtime_match.group(0))

    # -------------------------------------------------------------------------
    # STEP 4: If no special structure was detected, use the complete text.
    # -------------------------------------------------------------------------
    return normalize_text(text)


# =============================================================================
# 5. NORMALIZATION FOR FUZZY MATCHING
# =============================================================================

def normalize_for_matching(text):
    """
    Normalize dynamic values that should not dominate fuzzy similarity.

    Examples:
        "takes 1 positional argument"
        "takes 3 positional arguments"

    Both can be reduced to a common structural pattern.

    Variable names inside quotes are also normalized because the variable
    name itself is often not the important part of the error signature.

    Parameters
    ----------
    text : str
        Extracted error text.

    Returns
    -------
    str
        Text prepared for fuzzy comparison.
    """
    text = normalize_text(text)

    if not text:
        return ""

    # Replace integers and decimal numbers with <num>.
    text = re.sub(
        r"\b\d+(?:\.\d+)?\b",
        "<num>",
        text,
    )

    # Replace quoted variable names with <var>.
    text = re.sub(
        r"'[a-z_][a-z0-9_]*'",
        "<var>",
        text,
    )

    text = re.sub(
        r'"[a-z_][a-z0-9_]*"',
        "<var>",
        text,
    )

    # Collapse whitespace again after replacements.
    text = re.sub(r"\s+", " ", text).strip()

    return text


# =============================================================================
# 6. PREPARE REFERENCE BANK
# =============================================================================

def prepare_reference_bank():
    """
    Extract and normalize the matching text for every reference example.

    Returns
    -------
    list of dict
        Prepared reference examples.
    """
    prepared = []

    for reference in REFERENCE_EXAMPLES:
        ref = reference.copy()

        ref["clean_example"] = extract_error_text(ref["example"])
        ref["matching_text"] = normalize_for_matching(
            ref["clean_example"]
        )

        prepared.append(ref)

    return prepared


# =============================================================================
# 7. FUZZY SIMILARITY
# =============================================================================

def calculate_similarity(text1, text2):
    """
    Calculate the weighted fuzzy similarity between two texts.

    Components:
        - fuzz.ratio             -> 50%
        - fuzz.token_sort_ratio  -> 30%
        - fuzz.WRatio            -> 20%

    Parameters
    ----------
    text1, text2 : str
        Normalized strings to compare.

    Returns
    -------
    float
        Similarity score from 0 to 100.
    """
    if not text1 or not text2:
        return 0.0

    # Overall character-level similarity.
    score_ratio = fuzz.ratio(text1, text2)

    # Useful when the same words appear in a different order.
    score_token_sort = fuzz.token_sort_ratio(text1, text2)

    # General weighted fuzzy matching.
    score_wratio = fuzz.WRatio(text1, text2)

    # Weighted final score.
    score = (
        0.50 * score_ratio
        + 0.30 * score_token_sort
        + 0.20 * score_wratio
    )

    return round(score, 2)


# =============================================================================
# 8. CATEGORY-SPECIFIC CLASSIFIERS
# =============================================================================
#
# These functions make the four taxonomy categories explicit and reusable.
# The underlying algorithm remains the same: fuzzy matching against the
# manually audited references belonging to that category.
# =============================================================================

def _classify_against_references(matching_error, references):
    """
    Internal helper used by all four category-specific classifiers.

    It compares one normalized error against a supplied reference group,
    keeps the highest similarity score, and applies the global threshold.
    """
    best_score = 0.0
    best_reference = None

    for ref in references:
        score = calculate_similarity(
            matching_error,
            ref["matching_text"],
        )

        if score > best_score:
            best_score = score
            best_reference = ref

    if best_score > FUZZY_THRESHOLD:
        return best_reference, best_score

    return None, best_score


def classify_interface_formatting(matching_error, references):
    """
    Classify against Interface & Formatting references IF01–IF08.
    """
    category_references = [
        ref for ref in references
        if ref["taxonomy"] == "Interface & Formatting Errors"
    ]

    return _classify_against_references(
        matching_error,
        category_references,
    )


def classify_data_type_casting(matching_error, references):
    """
    Classify against Data Type & Casting references DT01–DT02.
    """
    category_references = [
        ref for ref in references
        if ref["taxonomy"] == "Data Type & Casting Errors"
    ]

    return _classify_against_references(
        matching_error,
        category_references,
    )


def classify_initialization_scope(matching_error, references):
    """
    Classify against Initialization & Scope references IS01–IS04.
    """
    category_references = [
        ref for ref in references
        if ref["taxonomy"] == "Initialization & Scope Errors"
    ]

    return _classify_against_references(
        matching_error,
        category_references,
    )


def classify_math_transformation(matching_error, references):
    """
    Classify against Math & Transformation references MT01–MT02.
    """
    category_references = [
        ref for ref in references
        if ref["taxonomy"] == "Math & Transformation Errors"
    ]

    return _classify_against_references(
        matching_error,
        category_references,
    )


# =============================================================================
# 9. MASTER CLASSIFIER
# =============================================================================

def classify_execution_feedback(feedback, references):
    """
    Classify one raw exec_feedback value.

    The actual classification compares the input against ALL manually
    audited references simultaneously. The taxonomy belonging to the
    highest-scoring reference becomes the predicted category.

    If the best score is not greater than FUZZY_THRESHOLD, the sample
    remains UNCLASSIFIED.

    Returns
    -------
    dict
        Classification result containing:
            - identified_error_type
            - matched_reference_id
            - fuzzy_similarity_score
            - matched_reference_example
    """
    # Handle missing feedback explicitly.
    if feedback is None or pd.isna(feedback):
        return {
            "identified_error_type": "NO_FEEDBACK",
            "matched_reference_id": "",
            "fuzzy_similarity_score": 0.0,
            "matched_reference_example": "",
        }

    if str(feedback).strip() == "":
        return {
            "identified_error_type": "NO_FEEDBACK",
            "matched_reference_id": "",
            "fuzzy_similarity_score": 0.0,
            "matched_reference_example": "",
        }

    # Extract the meaningful error message.
    clean_error = extract_error_text(feedback)

    # Normalize the error for structural comparison.
    matching_error = normalize_for_matching(clean_error)

    if not matching_error:
        return {
            "identified_error_type": "UNCLASSIFIED",
            "matched_reference_id": "",
            "fuzzy_similarity_score": 0.0,
            "matched_reference_example": "",
        }

    # Compare against every manually audited reference.
    best_score = -1.0
    best_reference = None

    for ref in references:
        score = calculate_similarity(
            matching_error,
            ref["matching_text"],
        )

        if score > best_score:
            best_score = score
            best_reference = ref

    # Assign the taxonomy only when the score exceeds 80%.
    if best_reference is not None and best_score > FUZZY_THRESHOLD:
        return {
            "identified_error_type": best_reference["taxonomy"],
            "matched_reference_id": best_reference["reference_id"],
            "fuzzy_similarity_score": best_score,
            "matched_reference_example": (
                best_reference["example"].strip()
            ),
        }

    # No forced classification.
    return {
        "identified_error_type": "UNCLASSIFIED",
        "matched_reference_id": (
            best_reference["reference_id"]
            if best_reference
            else ""
        ),
        "fuzzy_similarity_score": (
            best_score if best_score >= 0 else 0.0
        ),
        "matched_reference_example": (
            best_reference["example"].strip()
            if best_reference
            else ""
        ),
    }


# =============================================================================
# 10. REFERENCE BANK REPORT
# =============================================================================

def print_reference_summary(references):
    """Print a compact summary of the manually audited reference bank."""
    reference_df = pd.DataFrame(references)

    print("\n" + "=" * 80)
    print("MANUALLY AUDITED REFERENCE BANK")
    print("=" * 80)

    print(
        reference_df[
            ["reference_id", "taxonomy"]
        ].to_string(index=False)
    )

    print(f"\nTotal reference examples: {len(reference_df)}")

    print("\nReference examples per taxonomy:")
    print(reference_df["taxonomy"].value_counts())


# =============================================================================
# 11. SANITY TESTS
# =============================================================================

def run_sanity_tests(references):
    """
    Run representative examples before processing the complete dataset.

    These tests provide a quick check that the reference bank, extraction,
    normalization, and fuzzy matching pipeline are working as expected.
    """
    test_messages = [

        # Interface & Formatting
        """
        {'test_case': (0, 1, 0),
         'error': 'Runtime issue: TypeError: main() takes 1 positional argument but 3 were given'}
        """,

        # Interface & Formatting
        """
        {'test_case': ({}, {'key': 'value'}),
         'error': 'Runtime issue: TypeError: main() takes 0 positional arguments but 2 were given'}
        """,

        # Data Type & Casting
        """
        {'test_case': [1000, 2000],
         'error': "Runtime issue: TypeError: 'int' object is not iterable"}
        """,

        # Initialization & Scope
        """
        {'test_case': (-50, 5),
         'error': "Runtime issue: NameError: name 'a' is not defined"}
        """,

        # Initialization & Scope
        """
        {'test_case': (0, 0),
         'error': "Runtime issue: UnboundLocalError: cannot access local variable 'larger_number' where it is not associated with a value"}
        """,

        # Math & Transformation
        """
        {'test_case': (0, 0, 0, 1),
         'error': 'Runtime issue: ZeroDivisionError: division by zero'}
        """,

        # Math & Transformation
        """
        {'test_case': (0, 1, 1),
         'error': 'Runtime issue: ZeroDivisionError: float division by zero'}
        """,

        # Data Type & Casting
        """
        Test Case(-1, -1, 2, 1)
        Expected Output(0.666667, -0.333333)
        Actual Output : (0.666667, -0.33333333333333337)
        """,

        # Interface & Formatting
        """
        Test Case : (0, 0, 5, 5)
        Expected Output : (1.0, 0.0)
        Actual Output : ('1.000000', '0.000000')
        """,

        # Interface & Formatting
        """
        Test Case(0, 0, 1, 1)
        Expected Output(1.0, 0.0)
        Actual Output : {'slope': 1.0, 'intercept': 0.0}
        """,
    ]

    print("\n" + "=" * 80)
    print("MANUAL SANITY CHECK")
    print("=" * 80)

    for i, message in enumerate(test_messages, start=1):
        result = classify_execution_feedback(message, references)

        print(f"\nTest {i}")
        print("-" * 60)
        print("Classification:", result["identified_error_type"])
        print("Reference:", result["matched_reference_id"])
        print("Similarity:", result["fuzzy_similarity_score"])


# =============================================================================
# 12. CLASSIFY COMPLETE DATASET
# =============================================================================

def classify_dataset(df, references):
    """
    Classify every row in the input dataframe.

    Required column:
        exec_feedback

    Returns
    -------
    pandas.DataFrame
        Original data plus extracted/normalized feedback and classification
        results.
    """
    if "exec_feedback" not in df.columns:
        raise ValueError(
            "ERROR: The Excel file must contain a column named "
            "'exec_feedback'."
        )

    results = []
    total_rows = len(df)

    print("\n" + "=" * 80)
    print("CLASSIFYING COMPLETE DATASET")
    print("=" * 80)

    for index, feedback in enumerate(
        df["exec_feedback"],
        start=1,
    ):
        result = classify_execution_feedback(
            feedback,
            references,
        )

        results.append(result)

        # Display progress every 1,000 rows.
        if index % 1000 == 0:
            print(
                f"Processed {index:,} / {total_rows:,} rows"
            )

    results_df = pd.DataFrame(results)

    # Keep the original dataframe unchanged and add classification columns.
    df_result = df.copy()

    df_result["exec_feedback_clean"] = [
        extract_error_text(x)
        for x in df["exec_feedback"]
    ]

    df_result["exec_feedback_matching_text"] = [
        normalize_for_matching(
            extract_error_text(x)
        )
        for x in df["exec_feedback"]
    ]

    df_result["identified_error_type"] = (
        results_df["identified_error_type"]
    )

    df_result["matched_reference_id"] = (
        results_df["matched_reference_id"]
    )

    df_result["fuzzy_similarity_score"] = (
        results_df["fuzzy_similarity_score"]
    )

    df_result["matched_reference_example"] = (
        results_df["matched_reference_example"]
    )

    return df_result


# =============================================================================
# 13. CREATE SUMMARY
# =============================================================================

def create_summary(df_result):
    """
    Create error-type counts and percentages.
    """
    summary = (
        df_result["identified_error_type"]
        .value_counts()
        .reset_index()
    )

    summary.columns = [
        "Error Type",
        "Number of Rows",
    ]

    summary["Percentage"] = (
        summary["Number of Rows"]
        / len(df_result)
        * 100
    ).round(2)

    return summary


# =============================================================================
# 14. CREATE CATEGORY-SPECIFIC DATAFRAMES
# =============================================================================

def create_category_dataframes(df_result):
    """
    Split classified results into separate dataframes for each taxonomy.
    """
    categories = {
        "Interface & Formatting Errors": "interface_df",
        "Data Type & Casting Errors": "datatype_df",
        "Initialization & Scope Errors": "initialization_df",
        "Math & Transformation Errors": "math_df",
        "UNCLASSIFIED": "unclassified_df",
        "NO_FEEDBACK": "no_feedback_df",
    }

    output = {}

    for category, key in categories.items():
        output[key] = df_result[
            df_result["identified_error_type"] == category
        ].copy()

    return output


# =============================================================================
# 15. MATCH STATISTICS
# =============================================================================

def create_match_statistics(df_result, category_dfs):
    """
    Calculate overall classification and similarity statistics.
    """
    matched_df = df_result[
        ~df_result["identified_error_type"].isin(
            ["UNCLASSIFIED", "NO_FEEDBACK"]
        )
    ].copy()

    total_rows = len(df_result)
    classified_rows = len(matched_df)
    unclassified_rows = len(category_dfs["unclassified_df"])
    no_feedback_rows = len(category_dfs["no_feedback_df"])

    if classified_rows > 0:
        statistics = {
            "Total dataset rows": total_rows,
            "Rows classified": classified_rows,
            "Rows UNCLASSIFIED": unclassified_rows,
            "Rows with no feedback": no_feedback_rows,
            "Average similarity of classified rows": round(
                matched_df["fuzzy_similarity_score"].mean(),
                2,
            ),
            "Minimum similarity of classified rows": round(
                matched_df["fuzzy_similarity_score"].min(),
                2,
            ),
            "Maximum similarity of classified rows": round(
                matched_df["fuzzy_similarity_score"].max(),
                2,
            ),
        }
    else:
        statistics = {
            "Total dataset rows": total_rows,
            "Rows classified": 0,
            "Rows UNCLASSIFIED": unclassified_rows,
            "Rows with no feedback": no_feedback_rows,
        }

    return pd.DataFrame(
        {
            "Statistic": list(statistics.keys()),
            "Value": list(statistics.values()),
        }
    )


# =============================================================================
# 16. EXPORT RESULTS TO EXCEL
# =============================================================================

def export_results(
    df_result,
    summary,
    category_dfs,
    references,
    match_statistics,
    output_file,
):
    """
    Export the complete analysis to a multi-sheet Excel workbook.
    """
    reference_summary = pd.DataFrame(references)[
        [
            "reference_id",
            "taxonomy",
            "example",
        ]
    ].copy()

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
    ) as writer:

        # Complete classified dataset.
        df_result.to_excel(
            writer,
            sheet_name="All_Classified_Data",
            index=False,
        )

        # Overall summary.
        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )

        # Category-specific results.
        category_dfs["interface_df"].to_excel(
            writer,
            sheet_name="Interface_Formatting",
            index=False,
        )

        category_dfs["datatype_df"].to_excel(
            writer,
            sheet_name="Data_Type_Casting",
            index=False,
        )

        category_dfs["initialization_df"].to_excel(
            writer,
            sheet_name="Initialization_Scope",
            index=False,
        )

        category_dfs["math_df"].to_excel(
            writer,
            sheet_name="Math_Transformation",
            index=False,
        )

        # Samples requiring further manual audit.
        category_dfs["unclassified_df"].to_excel(
            writer,
            sheet_name="Unclassified_Audit",
            index=False,
        )

        # Rows without usable feedback.
        category_dfs["no_feedback_df"].to_excel(
            writer,
            sheet_name="No_Feedback",
            index=False,
        )

        # Reference bank used by the classifier.
        reference_summary.to_excel(
            writer,
            sheet_name="Reference_Summary",
            index=False,
        )

        # Similarity statistics.
        match_statistics.to_excel(
            writer,
            sheet_name="Match_Statistics",
            index=False,
        )


# =============================================================================
# 17. MAIN PROGRAM
# =============================================================================

def main():
    """
    Main execution pipeline.

    This function keeps the complete workflow organized:
        Load data
        -> prepare references
        -> sanity check
        -> classify
        -> summarize
        -> export
    """

    input_path = Path(INPUT_FILE)
    output_path = Path(OUTPUT_FILE)

    # -------------------------------------------------------------------------
    # Validate input file.
    # -------------------------------------------------------------------------
    if not input_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_path}\n"
            "Place the Excel dataset in the project directory or "
            "change INPUT_FILE at the top of this script."
        )

    print("=" * 80)
    print("EXECUTION ERROR TAXONOMY CLASSIFIER")
    print("=" * 80)

    # -------------------------------------------------------------------------
    # Load Excel dataset.
    # -------------------------------------------------------------------------
    print(f"\nReading input file: {input_path}")

    df = pd.read_excel(input_path)

    print("Dataset shape:", df.shape)

    print("\nColumns:")
    print(df.columns.tolist())

    # -------------------------------------------------------------------------
    # Validate required column.
    # -------------------------------------------------------------------------
    if "exec_feedback" not in df.columns:
        raise ValueError(
            "The Excel file must contain a column named 'exec_feedback'."
        )

    print("\n'exec_feedback' column found successfully.")

    # -------------------------------------------------------------------------
    # Prepare reference bank.
    # -------------------------------------------------------------------------
    references = prepare_reference_bank()

    print_reference_summary(references)

    # -------------------------------------------------------------------------
    # Run representative tests before processing the full dataset.
    # -------------------------------------------------------------------------
    run_sanity_tests(references)

    # -------------------------------------------------------------------------
    # Classify the complete dataset.
    # -------------------------------------------------------------------------
    df_result = classify_dataset(
        df,
        references,
    )

    # -------------------------------------------------------------------------
    # Generate summary.
    # -------------------------------------------------------------------------
    summary = create_summary(df_result)

    print("\n" + "=" * 80)
    print("FINAL CLASSIFICATION SUMMARY")
    print("=" * 80)

    print(summary.to_string(index=False))

    # -------------------------------------------------------------------------
    # Create category-specific dataframes.
    # -------------------------------------------------------------------------
    category_dfs = create_category_dataframes(df_result)

    # -------------------------------------------------------------------------
    # Calculate match statistics.
    # -------------------------------------------------------------------------
    match_statistics = create_match_statistics(
        df_result,
        category_dfs,
    )

    # -------------------------------------------------------------------------
    # Export all results.
    # -------------------------------------------------------------------------
    export_results(
        df_result=df_result,
        summary=summary,
        category_dfs=category_dfs,
        references=references,
        match_statistics=match_statistics,
        output_file=output_path,
    )

    # -------------------------------------------------------------------------
    # Final information.
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("PROCESSING COMPLETE")
    print("=" * 80)

    print("\nOutput file:", output_path)
    print("Total rows processed:", len(df_result))

    print(
        "Interface & Formatting:",
        len(category_dfs["interface_df"]),
    )

    print(
        "Data Type & Casting:",
        len(category_dfs["datatype_df"]),
    )

    print(
        "Initialization & Scope:",
        len(category_dfs["initialization_df"]),
    )

    print(
        "Math & Transformation:",
        len(category_dfs["math_df"]),
    )

    print(
        "UNCLASSIFIED:",
        len(category_dfs["unclassified_df"]),
    )

    print(
        "NO_FEEDBACK:",
        len(category_dfs["no_feedback_df"]),
    )


# =============================================================================
# 18. SCRIPT ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    main()
