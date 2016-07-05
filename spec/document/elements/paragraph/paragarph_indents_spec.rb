require 'spec_helper'

describe OoxmlParser::Indents do
  it 'indents_round' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/indents/indents_round.docx')
    expect(docx.element_by_description.first.ind.round(3).left_indent).to eq(1.25)
    expect(docx.element_by_description.first.ind.round(3).right_indent).to eq(0)
  end

  it 'indents_comparasion' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/indents/indents_comparasion.docx')
    expect(docx.element_by_description.first.ind.equal_with_round(OoxmlParser::Indents.new(0, 1.27, 0))).to be_truthy
  end
end
