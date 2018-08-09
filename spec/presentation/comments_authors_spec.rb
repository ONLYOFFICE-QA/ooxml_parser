require 'spec_helper'

describe 'CommentAuthors' do
  it 'simple_comment_authors' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/comment_authors/simple_comment_authors.pptx')
    expect(pptx.comment_authors.list.first.id).to eq(1)
    expect(pptx.comment_authors.list.first.name).to eq('Hamish Mitchell')
    expect(pptx.comment_authors.list.first.initials).to eq('HM')
    expect(pptx.comment_authors.list.first.last_index).to eq(1)
    expect(pptx.comment_authors.list.first.color_index).to eq(0)
  end
end
