require 'spec_helper'

describe 'Theme Color' do
  it 'Correct parse theeme color for 000 lastClr' do
    docx = OoxmlParser::Parser.parse('spec/document/theme_color/000_lastClr.docx')
    expect(docx.theme_colors.color_scheme[:dk1].color).to eq(OoxmlParser::Color.new(0, 0, 0))
  end

  it 'Do not crash if there is no theme.xml file' do
    docx = OoxmlParser::Parser.parse('spec/document/theme_color/no_theme_xml.docx')
    expect(docx.theme_colors).to be_nil
  end
end
