require 'spec_helper'

describe 'My behaviour' do
  it 'row_height_nil' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/row_height_nil.xlsx')
    expect(xlsx.worksheets.first.rows.first.height).to be_nil
  end

  it 'row_custom_height' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/row_custom_height.xlsx')
    expect(xlsx.worksheets.first.rows.first.custom_height).to be_truthy
  end

  it 'row_height in ooxml size' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/row_height_ooxml_size.xlsx')
    expect(xlsx.worksheets.first.rows.first.height).to eq(OoxmlParser::OoxmlSize.new(77.25, :point))
  end
end
