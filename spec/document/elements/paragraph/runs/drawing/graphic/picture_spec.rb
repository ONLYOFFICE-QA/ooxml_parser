# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocxPicture do
  it 'Check ImageWithInccorectLinkInRes' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/picture/image_with_inccorect_link.docx') }
      .to output("Cant find path to media file by id: rId6\n").to_stderr
  end

  it 'Check Image with empty link' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/picture/image_empty_link.docx')
    expect(docx).to be_with_data
  end
end
