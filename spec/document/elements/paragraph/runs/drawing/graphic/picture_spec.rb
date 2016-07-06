require 'spec_helper'

describe OoxmlParser::DocxPicture do
  it 'Check ImageWithInccorectLinkInRes' do
    expect { OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/graphic/picture/image_with_inccorect_link.docx') }.to raise_error(LoadError)
  end
end
