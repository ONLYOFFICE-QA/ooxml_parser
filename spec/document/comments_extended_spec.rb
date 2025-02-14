# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::CommentExtended do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('spec/document/comments_extended/comments_extended.docx') }

  it 'comments_extended paragraph id' do
    expect(docx.comments_extended[0].paragraph_id).to eq(1)
    expect(docx.comments_extended[1].paragraph_id).to eq(2)
  end

  it 'comments_extended parent_paragraph_id' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments_extended/comment_with_parent_id.docx')
    expect(docx.comments_extended[1].parent_paragraph_id).to eq(1)
  end

  it 'comments_extended done' do
    expect(docx.comments_extended[0].done).to be_truthy
    expect(docx.comments_extended[1].done).to be_falsey
  end

  it 'CommentsExtended#by_id returns nil for unknown id' do
    expect(docx.comments_extended.by_id(-1)).to be_nil
  end
end
