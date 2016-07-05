require 'spec_helper'

describe 'Runs Font Size' do
  it 'font_size_float' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/size/font_size_float.docx')
    expect(docx.elements[0].character_style_array[1].size).to eq(53.5)
  end
end
