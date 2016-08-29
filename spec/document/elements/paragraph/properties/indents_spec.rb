require 'spec_helper'

describe 'My behaviour' do
  it 'indents_default' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/properties/indents/indents_default.docx')
    expect(docx.elements[1].paragraph_properties.indent).to be_nil
  end
end
