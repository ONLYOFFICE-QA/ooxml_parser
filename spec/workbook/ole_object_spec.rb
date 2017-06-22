require 'spec_helper'

describe 'OleObject' do
  it 'ole_present.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/ole_object/ole_presnt.xlsx')
    expect(xlsx.worksheets.first.ole_objects.first).to be_a(OoxmlParser::AlternateContent)
  end
end
