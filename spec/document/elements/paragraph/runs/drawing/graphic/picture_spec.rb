require 'spec_helper'

describe OoxmlParser::DocxPicture do
  it 'Check ImageWithInccorectLinkInRes' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/picture/image_with_inccorect_link.docx') }
      .to output("Cant find path to media file by id: rId6\n").to_stderr
  end
end
