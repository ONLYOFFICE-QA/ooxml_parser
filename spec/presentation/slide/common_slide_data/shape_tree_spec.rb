require 'spec_helper'

describe 'Connection Shape' do
  it 'shape_tree.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/common_slide_data/shape_tree/shape_tree.pptx')
    expect(pptx.slides.first.common_slide_data.shape_tree.elements.length).to eq(2)
  end
end
