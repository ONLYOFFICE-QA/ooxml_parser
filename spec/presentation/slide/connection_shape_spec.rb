require 'spec_helper'

describe 'Connection Shape' do
  it 'connection_shape_exists.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/connection_shape/connection_shape_exists.pptx')
    expect(pptx.slides.first.elements.first).to be_a(OoxmlParser::ConnectionShape)
  end
end
