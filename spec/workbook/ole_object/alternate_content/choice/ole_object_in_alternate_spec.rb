require 'spec_helper'

describe 'OleObjectInAlternateContent' do
  it 'ole_object_in_alternate' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/ole_object/alternate_content/choice/ole_object/ole_object_in_alternate.xlsx')
    expect(xlsx.worksheets.first.ole_objects.first.office2010_content.ole_object).to be_a(OoxmlParser::OleObject)
  end
end
