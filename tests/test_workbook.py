from recon.identifiers import clean, is_valid_job_no, search_terms

def test_clean_strips_prefix():
    assert clean("RFQ 6100175731") == "6100175731"
    assert clean("NA") is None

def test_job_no_format():
    assert is_valid_job_no("7373260")
    assert not is_valid_job_no("77372045")

def test_search_terms_dedupes():
    t = search_terms({"job_no": "7365554", "rfq": "SQ850584", "order_ref": "SQ850584"})
    assert [v for _, v in t] == ["7365554", "SQ850584"]
