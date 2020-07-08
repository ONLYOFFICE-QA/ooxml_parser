# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::OOXMLDocumentObject do
  describe '#==' do
    it 'Compare two inherited classes without any instance variables' do
      expect(OoxmlParser::BookmarkStart.new).not_to eq(OoxmlParser::BookmarkEnd.new)
    end
  end
end
