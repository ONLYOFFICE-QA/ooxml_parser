require 'spec_helper'

describe 'My behaviour' do
  it 'tranistion_sound_action' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/transition/sound_action/tranistion_sound_action.pptx')
    expect(pptx.slides.first.transition.sound_action.start_sound.file_reference.content.length).to be > 1000
  end
end
