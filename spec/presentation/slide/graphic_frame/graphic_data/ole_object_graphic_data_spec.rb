# frozen_string_literal: true

require 'spec_helper'

describe 'GraphicFrame - OleObject' do
  it 'GraphicData - OleObject' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/ole_object/graphic_data_ole.pptx')
    expect(pptx.slides.last.elements.last.graphic_data.first.file_reference.content).not_to be_empty
  end
end
