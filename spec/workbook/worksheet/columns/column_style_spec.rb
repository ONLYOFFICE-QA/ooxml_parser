require 'spec_helper'

describe 'My behaviour' do
  it 'column_style.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/columns/style/column_style.xlsx')
    expect(xlsx.worksheets.first.columns.first.style.fill_color.pattern_fill.background_color).to eq(OoxmlParser::Color.new(91, 155, 213))
  end
end
