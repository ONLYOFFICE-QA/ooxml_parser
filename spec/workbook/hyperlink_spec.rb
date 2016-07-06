require 'spec_helper'

describe 'My behaviour' do
  it 'hyperlink_internal.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/hyperlinks/hyperlink_internal.xlsx')
    expect(xlsx.worksheets[0].hyperlinks[0].link).to eq(OoxmlParser::Coordinates.new(10, 'F'))
    expect(xlsx.worksheets[0].rows[0].cells.first.text).to eq('yandex')
    expect(xlsx.worksheets[0].hyperlinks[0].tooltip).to eq('go to www.yandex.ru')
  end
end
