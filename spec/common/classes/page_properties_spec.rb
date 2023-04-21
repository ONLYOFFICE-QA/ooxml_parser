# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::PageProperties do
  describe '#section_break' do
    it 'returns "Odd page" when type is "oddPage"' do
      page_properties = described_class.new(type: 'oddPage')
      expect(page_properties.section_break).to eq('Odd page')
    end

    it 'returns "Current Page" when type is "continuous"' do
      page_properties = described_class.new(type: 'continuous')
      expect(page_properties.section_break).to eq('Current Page')
    end

    it 'returns "Next Page" when type is anything else' do
      page_properties = described_class.new(type: 'fake-type')
      expect(page_properties.section_break).to eq('Next Page')
    end
  end
end
