require 'spec_helper'

describe 'Font color' do
  it 'font_color.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/styles/fonts/font_color/font_color.xlsx')
    expect(xlsx.style_sheet.fonts.fonts_array[1].color).to eq(OoxmlParser::Color.new(217, 217, 217))
  end
end
