# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OOXMLDocumentObject do
  describe '#==' do
    it 'Compare two inherited classes without any instance variables' do
      expect(OoxmlParser::BookmarkStart.new).not_to eq(OoxmlParser::BookmarkEnd.new)
    end
  end

  describe 'Class methods' do
    describe 'OOXMLDocumentObject.encrypted_file?' do
      it 'OOXMLDocumentObject.encrypted_file?(path, ignore_system: true) always return false' do
        expect(described_class).not_to be_encrypted_file(__FILE__, ignore_system: true)
      end
    end
  end
end
