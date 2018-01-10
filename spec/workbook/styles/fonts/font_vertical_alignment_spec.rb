require 'spec_helper'

describe 'Font Vertical Alignment' do
  it 'font_vertical_alignment.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/styles/fonts/font_vertical_alignment/font_vertical_alignment.xlsx')
    expect(xlsx.style_sheet.fonts
               .fonts_array[1]
               .vertical_alignment.value).to eq(:superscript)
  end
end
