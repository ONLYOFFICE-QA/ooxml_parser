require 'spec_helper'

describe 'My behaviour' do
  it 'indents_default' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/properties/indents/indents_default.docx')
    expect(docx.elements[1].paragraph_properties.indent).to be_nil
  end

  it 'indents_custom' do
    pptx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/properties/indents/indents_custom.pptx')
    expect(pptx.slides[0].elements.last.text_body.paragraphs[0].properties.indent).to eq(OoxmlParser::OoxmlSize.new(1.02, :centimeter))
  end
end
