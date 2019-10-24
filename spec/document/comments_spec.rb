# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Comment do
  it 'comment_object.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments/comment_object.docx')
    expect(docx.element_by_description(location: :comment, type: :paragraph)[0].character_style_array[0].text).to eq('Is it true?')
  end

  it 'comments_paragraphs_with_ids.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments/comments_paragraphs_with_ids.docx')
    expect(docx.comments[0].paragraphs[0].paragraph_id).to eq(1)
    expect(docx.comments[0].paragraphs[0].text_id).to eq(1)
  end

  it 'Comment Extended data' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments/comments_paragraphs_with_ids.docx')
    expect(docx.comments[0].paragraphs[0].comment_extend_data.done).to be_truthy
  end
end
