require 'spec_helper'

describe 'Workbook#with_data' do
  it 'no_data' do
    xlsx = OoxmlParser::Parser.parse('spec/workbook/chartsheets/file_with_chartsheets.xlsx')
    expect(xlsx.worksheets.length).to eq(3)
  end
end
