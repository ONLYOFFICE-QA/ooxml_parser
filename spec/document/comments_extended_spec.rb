# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::CommentExtended do
  it 'comments_extended paragraph id' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments_extended/comments_extended.docx')
    expect(docx.comments_extended[0].paragraph_id).to eq(1)
    expect(docx.comments_extended[1].paragraph_id).to eq(2)
  end

  it 'comments_extended done' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments_extended/comments_extended.docx')
    expect(docx.comments_extended[0].done).to be_truthy
    expect(docx.comments_extended[1].done).to be_falsey
  end
end
