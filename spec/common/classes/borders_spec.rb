# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Borders do
  let(:default_borders) { OoxmlParser::Borders.new }

  describe 'Borders#copy' do
    it 'Borders#copy result same as original' do
      expect(default_borders).to eq(default_borders.copy)
    end
  end

  describe 'Borders#each_border' do
    it 'Borders#each_border return BordersProperties' do
      default_borders.each_border do |border|
        expect(border).to be_a(OoxmlParser::BordersProperties)
      end
    end
  end

  describe 'Borders#to_s' do
    it 'Borders#to_s return data' do
      expect(default_borders.to_s).to include('Left border')
    end
  end

  describe 'Borders#visible?' do
    it 'Borders#visible? is false for default borders' do
      expect(default_borders).not_to be_visible
    end
  end
end
