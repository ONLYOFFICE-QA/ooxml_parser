require 'spec_helper'

describe 'My behaviour' do
  it 'comments.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/comments/comments.xlsx')
    expect(xlsx.worksheets.first.comments.comments.first.characters.first.text).to eq('Hello')
  end

  it 'comment_text_object.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/comments/comment_text_object.xlsx')
    expect(xlsx.worksheets.first.comments.comment_list.comments.first.text.text).to include('Hi')
  end
end
