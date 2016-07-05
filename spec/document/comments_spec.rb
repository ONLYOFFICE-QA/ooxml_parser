require 'spec_helper'

describe OoxmlParser::Comment do
  it 'comment_object.docx' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/comments/comment_object.docx')
    expect(docx.element_by_description(location: :comment, type: :paragraph)[0].character_style_array[0].text).to eq('Is it true?')
  end
end
