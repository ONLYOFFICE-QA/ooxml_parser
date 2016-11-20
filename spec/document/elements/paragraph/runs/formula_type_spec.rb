require 'spec_helper'

describe 'My behaviour' do
  it 'pre_sub_superscrip_formula' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/formula_run/pre_sub_superscrip_formula.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::DocxFormula)
  end
end
