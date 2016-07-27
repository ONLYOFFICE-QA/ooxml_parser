require 'spec_helper'

describe 'Runs Font Size' do
  it 'font_size_float' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/size/font_size_float.docx')
    expect(docx.elements[0].character_style_array[1].size).to eq(53.5)
  end

  it 'default_font_size' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/size/default_font_size.docx')
    expect(docx.elements.first.character_style_array[1].size).to eq(15)
  end
end
