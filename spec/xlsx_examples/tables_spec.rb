require 'rspec'
require 'ooxml_parser'

describe 'Check table.xml file' do
  it 'relation_to_table' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/xlsx_examples/tables/relation_to_table.xlsx')
    expect(xlsx.worksheets.first.relationships[0].target).to eq('../tables/table1.xml')
  end
end
