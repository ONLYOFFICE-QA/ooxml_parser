require 'spec_helper'

describe 'My behaviour' do
  it 'row_height_nil' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/rows/row_height_nil.xlsx')
    expect(xlsx.worksheets.first.rows.first.height).to be_nil
  end
end
