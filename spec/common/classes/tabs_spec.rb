# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Tabs do
  let(:tabs) { described_class.new }

  describe '#empty?' do
    it '#empty? is true for empty class' do
      expect(tabs).to be_empty
    end

    it '#empty? is false for non-empty class' do
      tabs.tabs_array = ['foo']
      expect(tabs).not_to be_empty
    end
  end
end
