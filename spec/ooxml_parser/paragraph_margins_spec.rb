# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ParagraphMargins do
  let(:some_size) { OoxmlParser::OoxmlSize.new(0, :centimeter) }

  describe 'ParagraphMargins side initialization' do
    it 'ParagraphMargins#top can be initialized in constructor' do
      expect(described_class.new(some_size).top).to eq(some_size)
    end

    it 'ParagraphMargins#bottom can be initialized in constructor' do
      expect(described_class.new(nil, some_size).bottom).to eq(some_size)
    end

    it 'ParagraphMargins#left can be initialized in constructor' do
      expect(described_class.new(nil, nil, some_size).left).to eq(some_size)
    end

    it 'ParagraphMargins#right can be initialized in constructor' do
      expect(described_class.new(nil, nil, nil, some_size).right).to eq(some_size)
    end
  end
end
