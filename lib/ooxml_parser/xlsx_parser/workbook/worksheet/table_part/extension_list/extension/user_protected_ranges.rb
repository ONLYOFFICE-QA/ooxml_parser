# frozen_string_literal: true

module OoxmlParser
  # Class for `UserProtectedRanges` data
  class UserProtectedRanges < OOXMLDocumentObject
    # @return [Array<ProtectedRange>] list of UserProtectedRange grodateup
    attr_reader :user_protected_ranges

    def initialize(parent: nil)
      @user_protected_ranges = []
      super
    end

    # @return [UserProtectedRange] accessor
    def [](key)
      user_protected_ranges[key]
    end

    # Parse UserProtectedRanges data
    # @param [Nokogiri::XML:Element] node with UserProtectedRanges data
    # @return [UserProtectedRange] value of UserProtectedRanges data
    def parse(node)
      node.xpath('*').each do |range_node|
        case range_node.name
        when 'userProtectedRange'
          @user_protected_ranges << ProtectedRange.new(parent: self).parse(range_node)
        end
      end
      self
    end
  end
end
