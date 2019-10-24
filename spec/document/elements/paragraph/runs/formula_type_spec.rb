# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'pre_sub_superscrip_formula' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/formula_run/pre_sub_superscrip_formula.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end

  it 'delimeter_begin_char_empty' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/formula_run/delimeter_begin_char_empty.docx')
    expect(docx).to be_with_data
  end
end
