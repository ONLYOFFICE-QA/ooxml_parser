require 'spec_helper'

describe 'Notes' do
  it 'slide_raltionship.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/notes/notes.pptx')
    expect(pptx.slides.first.relationships[0]).to be_a(OoxmlParser::Relationship)
  end

  it 'slide_notes.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/notes/notes.pptx')
    expect(pptx.slides.first.note.common_slide_data).to be_a(OoxmlParser::CommonSlideData)
  end
end
