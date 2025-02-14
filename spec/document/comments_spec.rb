# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Comment do
  let(:docx) { OoxmlParser::DocxParser.parse_docx('/home/serov/sources/ooxml_parser/spec/document/comments/comment_object.docx') }

  it 'comment_object.docx' do
    expect(docx.element_by_description(location: :comment, type: :paragraph)[0].character_style_array[0].text).to eq('Is it true?')
  end

  describe 'Comment additional params' do
    it 'comment object have auth' do
      expect(docx.comments[0].author).to eq('Hamish Mitchell')
    end

    it 'comment object have date in string format' do
      expect(docx.comments[0].date_string).to eq('2015-04-06T13:09:31Z')
    end

    it 'comment object have initials' do
      expect(docx.comments[0].initials).to eq('HM')
    end
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
