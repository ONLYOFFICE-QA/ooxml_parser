# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::HeaderFooterChild do
  describe 'Header has only left part' do
    let(:header) { described_class.new(raw_string: '&L&"Arial"&14Confidential') }

    it 'Has left part' do
      expect(header.left).to eq('&"Arial"&14Confidential')
    end

    it 'Center part is nil' do
      expect(header.center).to be_nil
    end

    it 'Right part is nil' do
      expect(header.right).to be_nil
    end
  end

  describe 'Header has only center part' do
    let(:header) { described_class.new(raw_string: '&C&"-,Bold"&Kff0000&D') }

    it 'Left part is nil' do
      expect(header.left).to be_nil
    end

    it 'Has center part' do
      expect(header.center).to eq('&"-,Bold"&Kff0000&D')
    end

    it 'Right part is nil' do
      expect(header.right).to be_nil
    end
  end

  describe 'Header has only right part' do
    let(:header) { described_class.new(raw_string: '&R&"-,Italic"&UPage &P') }

    it 'Left part is nil' do
      expect(header.left).to be_nil
    end

    it 'Center part is nil' do
      expect(header.center).to be_nil
    end

    it 'Has right part' do
      expect(header.right).to eq('&"-,Italic"&UPage &P')
    end
  end

  describe 'Header does not have left part' do
    let(:header) { described_class.new(raw_string: '&C&"-,Bold"&Kff0000&D&R&"-,Italic"&UPage &P') }

    it 'Left part is nil' do
      expect(header.left).to be_nil
    end

    it 'Has center part' do
      expect(header.center).to eq('&"-,Bold"&Kff0000&D')
    end

    it 'Has right part' do
      expect(header.right).to eq('&"-,Italic"&UPage &P')
    end
  end

  describe 'Header does not have center part' do
    let(:header) { described_class.new(raw_string: '&L&"Arial"&14Confidential&R&"-,Italic"&UPage &P') }

    it 'Has Left part' do
      expect(header.left).to eq('&"Arial"&14Confidential')
    end

    it 'Center part is nil' do
      expect(header.center).to be_nil
    end

    it 'Has right part' do
      expect(header.right).to eq('&"-,Italic"&UPage &P')
    end
  end

  describe 'Header does not have right part' do
    let(:header) { described_class.new(raw_string: '&L&"Arial"&14Confidential&C&"-,Bold"&Kff0000&D') }

    it 'Has Left part' do
      expect(header.left).to eq('&"Arial"&14Confidential')
    end

    it 'Has center part' do
      expect(header.center).to eq('&"-,Bold"&Kff0000&D')
    end

    it 'Right part is nil' do
      expect(header.right).to be_nil
    end
  end
end
