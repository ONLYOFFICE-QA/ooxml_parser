# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <sheetProtection> tag
  class SheetProtection < OOXMLDocumentObject
    # @return [String] Name of hashing algorithm
    attr_reader :algorithm_name
    # @return [String] Hash value for the password
    attr_reader :hash_value
    # @return [String] Salt value for the password
    attr_reader :salt_value
    # @return [Integer] Number of times the hashing function shall be iteratively run
    attr_reader :spin_count
    # @return [True, False] Specifies if sheet is protected
    attr_reader :sheet
    # @return [True, False] Specifies if using autofilter is not allowed on protected sheet
    attr_reader :auto_filter
    # @return [True, False] Specifies if deleting columns is not allowed on protected sheet
    attr_reader :delete_columns
    # @return [True, False] Specifies if deleting rows is not allowed on protected sheet
    attr_reader :delete_rows
    # @return [True, False] Specifies if formatting cells is not allowed on protected sheet
    attr_reader :format_cells
    # @return [True, False] Specifies if formatting columns is not allowed on protected sheet
    attr_reader :format_columns
    # @return [True, False] Specifies if formatting rows is not allowed on protected sheet
    attr_reader :format_rows
    # @return [True, False] Specifies if inserting columns is not allowed on protected sheet
    attr_reader :insert_columns
    # @return [True, False] Specifies if inserting rows is not allowed on protected sheet
    attr_reader :insert_rows
    # @return [True, False] Specifies if inserting hyperlinks is not allowed on protected sheet
    attr_reader :insert_hyperlinks
    # @return [True, False] Specifies if editing objects is not allowed on protected sheet
    attr_reader :objects
    # @return [True, False] Specifies if using pivot tables is not allowed on protected sheet
    attr_reader :pivot_tables
    # @return [True, False] Specifies if editing scenarios is not allowed on protected sheet
    attr_reader :scenarios
    # @return [True, False] Specifies if selecting locked cells is not allowed on protected sheet
    attr_reader :select_locked_cells
    # @return [True, False] Specifies if selecting unlocked cells is not allowed on protected sheet
    attr_reader :select_unlocked_cells
    # @return [True, False] Specifies if sorting is not allowed on protected sheet
    attr_reader :sort

    def initialize(parent: nil)
      @objects = false
      @scenarios = false
      @select_locked_cells = false
      @select_unlocked_cells = false
      @auto_filter = true
      @delete_columns = true
      @delete_rows = true
      @format_cells = true
      @format_columns = true
      @format_rows = true
      @insert_columns = true
      @insert_rows = true
      @insert_hyperlinks = true
      @pivot_tables = true
      @sort = true
      super
    end

    # Parse SheetProtection data
    # @param [Nokogiri::XML:Element] node with SheetProtection data
    # @return [Sheet] value of SheetProtection
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'algorithmName'
          @algorithm_name = value.value.to_s
        when 'hashValue'
          @hash_value = value.value.to_s
        when 'saltValue'
          @salt_value = value.value.to_s
        when 'spinCount'
          @spin_count = value.value.to_i
        when 'sheet'
          @sheet = attribute_enabled?(value)
        when 'autoFilter'
          @auto_filter = attribute_enabled?(value)
        when 'deleteColumns'
          @delete_columns = attribute_enabled?(value)
        when 'deleteRows'
          @delete_rows = attribute_enabled?(value)
        when 'formatCells'
          @format_cells = attribute_enabled?(value)
        when 'formatColumns'
          @format_columns = attribute_enabled?(value)
        when 'formatRows'
          @format_rows = attribute_enabled?(value)
        when 'insertColumns'
          @insert_columns = attribute_enabled?(value)
        when 'insertRows'
          @insert_rows = attribute_enabled?(value)
        when 'insertHyperlinks'
          @insert_hyperlinks = attribute_enabled?(value)
        when 'objects'
          @objects = attribute_enabled?(value)
        when 'pivotTables'
          @pivot_tables = attribute_enabled?(value)
        when 'scenarios'
          @scenarios = attribute_enabled?(value)
        when 'selectLockedCells'
          @select_locked_cells = attribute_enabled?(value)
        when 'selectUnlockedCells'
          @select_unlocked_cells = attribute_enabled?(value)
        when 'sort'
          @sort = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
