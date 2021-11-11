# frozen_string_literal: true

require 'spec_helper'

describe 'Transition#sound_action' do
  it 'tranistion_sound_action' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/transition/sound_action/tranistion_sound_action.pptx')
    expect(pptx.slides.first.transition.sound_action.start_sound.file_reference.content.length).to be > 1000
  end

  it 'sound_action_without_sound_name' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/transition/sound_action/transition_without_sound_name.pptx')
    expect(pptx.slides.first.transition
               .sound_action.start_sound.name).to be_empty
  end
end
