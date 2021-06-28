# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::XlsxHeaderFooter do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/header_footer/header_footer.xlsx') }
  let(:header_footer) { xlsx.worksheets.first.header_footer }

  it 'Has align_with_margins' do
    expect(header_footer.align_with_margins).to be_falsey
  end

  it 'Has different_first' do
    expect(header_footer.different_first).to be_truthy
  end

  it 'Has different_odd_even' do
    expect(header_footer.different_odd_even).to be_truthy
  end

  it 'Has scale_with_document' do
    expect(header_footer.scale_with_document).to be_falsey
  end
end
