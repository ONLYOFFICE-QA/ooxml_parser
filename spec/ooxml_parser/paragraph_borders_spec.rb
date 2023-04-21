# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ParagraphBorders do
  let(:empty_border) { described_class.new }
  let(:single_line_border) do
    OoxmlParser::BordersProperties.new(:auto,
                                       1,
                                       :single)
  end

  describe '#border_visual_type' do
    before do
      empty_border.left = OoxmlParser::BordersProperties.new
      empty_border.right = OoxmlParser::BordersProperties.new
      empty_border.top = OoxmlParser::BordersProperties.new
      empty_border.bottom = OoxmlParser::BordersProperties.new
      empty_border.between = OoxmlParser::BordersProperties.new
    end

    it 'none border_visual_type' do
      expect(empty_border.border_visual_type).to eq(:none)
    end

    it 'left border_visual_type' do
      empty_border.left = single_line_border
      expect(empty_border.border_visual_type).to eq(:left)
    end

    it 'right border_visual_type' do
      empty_border.right = single_line_border
      expect(empty_border.border_visual_type).to eq(:right)
    end

    it 'top border_visual_type' do
      empty_border.top = single_line_border
      expect(empty_border.border_visual_type).to eq(:top)
    end

    it 'bottom border_visual_type' do
      empty_border.bottom = single_line_border
      expect(empty_border.border_visual_type).to eq(:bottom)
    end

    it 'inner border_visual_type' do
      empty_border.between = single_line_border
      expect(empty_border.border_visual_type).to eq(:inner)
    end

    it 'all border_visual_type' do
      empty_border.top = single_line_border
      empty_border.bottom = single_line_border
      empty_border.left = single_line_border
      empty_border.right = single_line_border
      empty_border.between = single_line_border
      expect(empty_border.border_visual_type).to eq(:all)
    end

    it 'outer border_visual_type' do
      empty_border.top = single_line_border
      empty_border.bottom = single_line_border
      empty_border.left = single_line_border
      empty_border.right = single_line_border
      expect(empty_border.border_visual_type).to eq(:outer)
    end
  end
end
