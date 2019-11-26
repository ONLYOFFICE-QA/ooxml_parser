# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocxParagraphRun do
  let(:run) { described_class.new }

  describe 'equality' do
    it 'Run is equal to itself' do
      expect(run).to eq(run)
    end
  end
end
