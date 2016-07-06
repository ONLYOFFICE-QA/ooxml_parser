require 'spec_helper'

describe 'My behaviour' do
  it 'DefaultFontName' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/font_name/default_font_name.pptx')
    expect(pptx.slides.first.nonempty_elements.first.text_body.paragraphs[0].characters[0].properties.font_name).to eq('Arial')
  end
end
