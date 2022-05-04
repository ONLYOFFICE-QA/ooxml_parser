# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <workbookProtection> tag
  class WorkbookProtection < OOXMLDocumentObject
    # @return [True, False] Specifies if workbook structure is protected
    attr_reader :lock_structure
    # @return [String] name of hashing algorithm
    attr_reader :workbook_algorithm_name
    # @return [String] hash value for the password
    attr_reader :workbook_hash_value
    # @return [String] salt value for the password
    attr_reader :workbook_salt_value
    # @return [Integer] number of times the hashing function shall be iteratively run
    attr_reader :workbook_spin_count

    # Parse WorkbookProtection data
    # @param [Nokogiri::XML:Element] node with WorkbookProtection data
    # @return [Sheet] value of WorkbookProtection
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'lockStructure'
          @lock_structure = attribute_enabled?(value)
        when 'workbookAlgorithmName'
          @workbook_algorithm_name = value.value.to_s
        when 'workbookHashValue'
          @workbook_hash_value = value.value.to_s
        when 'workbookSaltValue'
          @workbook_salt_value = value.value.to_s
        when 'workbookSpinCount'
          @workbook_spin_count = value.value.to_i
        end
      end
      self
    end
  end
end
