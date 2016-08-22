require 'spec_helper'

describe 'My behaviour' do
  it 'border_color_nil' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/properties/borders/border_color_nil.docx')
    expect(docx.elements.first.paragraph_properties.paragraph_borders.left).to be_nil
    expect(docx.elements.first.paragraph_properties.paragraph_borders.right.color).not_to be_nil
  end
end
