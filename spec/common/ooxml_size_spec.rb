# frozen_string_literal: true

require 'spec_helper'

describe 'OoxmlSize' do
  it 'Twips and dxa units should be the same' do
    expect(OoxmlParser::OoxmlSize.new(100, :dxa)).to eq(OoxmlParser::OoxmlSize.new(100, :twip))
  end

  describe 'one_100th_point' do
    it '1 point converts to one_100th_point' do
      expect(OoxmlParser::OoxmlSize.new(1, :point).to_unit(:one_100th_point)).to eq(OoxmlParser::OoxmlSize.new(100, :one_100th_point))
    end
  end

  describe 'one_1000th_percent' do
    it 'compare same - percent and one_1000th_percent' do
      expect(OoxmlParser::OoxmlSize.new(44, :percent)).to eq(OoxmlParser::OoxmlSize.new(43_999, :one_1000th_percent))
    end

    it 'compare different - percent and one_1000th_percent' do
      expect(OoxmlParser::OoxmlSize.new(43, :percent)).not_to eq(OoxmlParser::OoxmlSize.new(43_999, :one_1000th_percent))
    end
  end

  describe 'pct' do
    it 'compare same - percent and one 50th of percent' do
      expect(OoxmlParser::OoxmlSize.new(44, :percent)).to eq(OoxmlParser::OoxmlSize.new(44 * 50, :pct))
    end

    it 'compare different - percent and one 50th of percent' do
      expect(OoxmlParser::OoxmlSize.new(43, :percent)).not_to eq(OoxmlParser::OoxmlSize.new(43.99 * 50, :pct))
    end
  end
end
