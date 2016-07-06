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
                   .common_time_node.children.first.common_time_node.children.first.common_time_node.start_conditions.first.event).to be_nil
      end
    end
  end
end
