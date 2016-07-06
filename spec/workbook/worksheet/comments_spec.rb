require 'spec_helper'

describe 'My behaviour' do
  it 'comments.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/comments/comments.xlsx')
    expect(xlsx.worksheets.first.comments.comments.first.characters.first.text).to eq('Hello')
  end
end
