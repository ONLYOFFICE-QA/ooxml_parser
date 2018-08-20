require 'spec_helper'

describe 'CommentsRange' do
  it 'CommentRangeStart is not empty' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/comments_range/comments_range.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::CommentRangeStart)
  end

  it 'CommentRangeStart has method to get comment' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/comments_range/comments_range.docx')
    expect(docx.elements.first.nonempty_runs.first.comment.paragraphs.first.nonempty_runs.first.text).to eq('Comment')
  end
end
