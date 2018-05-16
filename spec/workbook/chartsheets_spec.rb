require 'spec_helper'

describe 'Workbook#with_data' do
  it 'no_data' do
    xlsx = OoxmlParser::Parser.parse('/home/lobashov/Downloads/Telegram Desktop/TestCase_1346154671.284118.docx')
    expect(xlsx.worksheets.length).to eq(3)
  end
end
