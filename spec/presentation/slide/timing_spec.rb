# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  describe 'time_node' do
    describe 'animation' do
      it 'animation_no_transition' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/timing/time_node/animation/animation_no_transition.pptx')
        expect(pptx.slides[10].timing.time_node_list.first.common_time_node.children.first
                   .common_time_node.children.first.common_time_node.children.first.common_time_node
                   .children.first.common_time_node.children[1].transition).to be_nil
      end
    end

    describe 'condition' do
      it 'condition_no_event' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/timing/time_node/condition/condition_no_event.pptx')
        expect(pptx.slides[1].timing.time_node_list.first.common_time_node.children.first
                   .common_time_node.children.first.common_time_node.children.first.common_time_node.start_conditions.list.first.event).to be_nil
      end

      it 'condition_list_check' do
        pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/timing/time_node/condition/condition_list_check.pptx')
        expect(pptx.slides.first.timing.time_node_list.first.common_time_node.children.first
                   .common_time_node.children.first.common_time_node.start_conditions.list.first.delay).to eq(:indefinite)
      end
    end

    it 'Audio time node' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/' \
                                                'timing/time_node/' \
                                                'time_node_audio.pptx')
      expect(pptx.slides.first.timing
                 .time_node_list.first
                 .common_time_node
                 .children[1].type).to eq(:audio)
    end
  end
end
