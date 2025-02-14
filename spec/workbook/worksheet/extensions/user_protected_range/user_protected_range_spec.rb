# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::UserProtectedRanges do
  let(:xlsx) { OoxmlParser::Parser.parse('spec/workbook/worksheet/extensions/user_protected_range/user_protected_range.xlsx') }
  let(:user_protected_ranges) { xlsx.worksheets[0].extension_list[0].user_protected_ranges }

  it 'Has user_protected_ranges' do
    expect(user_protected_ranges.user_protected_ranges.size).to eq(1)
  end

  it 'Has name' do
    expect(user_protected_ranges[0].name).to eq('protectedRange')
  end

  it 'Has reference_sequence' do
    expect(user_protected_ranges[0].reference_sequence).to eq('$A$1:$B$1')
  end
end
