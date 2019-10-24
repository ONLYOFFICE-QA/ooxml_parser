# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Cell Text Not Empty' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/text/text_not_empty.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].text).not_to be_empty
  end

  it 'ABS_emb' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/text/abs_formula.xlsx')
    expect(xlsx.worksheets.first.rows.first.cells.first.text).to eq('Data')
  end

  it 'LOOKUP_emb' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/text/lookup_formula.xlsx')
    expect(xlsx.worksheets[1].rows[5].cells.first.text).to be_empty
  end

  it 'FormulaDoc' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/text/formula_doc.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('a')
  end

  it 'CellTextNotStartWithQuote' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/text/text_not_starting_with_quote.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq('abc')
    expect(xlsx.worksheets[0].rows[0].cells[0].raw_text).to eq('abc')
  end

  it 'CellTextStartWithQuote' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/cells/text/text_starting_with_quote.xlsx')
    expect(xlsx.worksheets[0].rows[0].cells[0].text).to eq("'abc")
    expect(xlsx.worksheets[0].rows[0].cells[0].raw_text).to eq('abc')
  end
end
