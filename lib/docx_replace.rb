# encoding: UTF-8

require "docx_replace/version"
require 'zip'
require 'tempfile'

module DocxReplace
  class Doc
    attr_reader :document_content

    def initialize(path, temp_dir=nil, files = [])
      @zip_file = Zip::File.new(path)
      @temp_dir = temp_dir
      files.each do |file|
        read_docx_file(file)
      end
    end

    def replace(pattern, replacement, multiple_occurrences=false)
      replace = replacement.to_s.encode(xml: :text)
       @document_content.each do |key, doc|     
        Rails.logger.info "Scanning: #{key}"
        if multiple_occurrences            
          @document_content[key].force_encoding("UTF-8").gsub!(pattern, replace)
        else
          @document_content[key].force_encoding("UTF-8").sub!(pattern, replace)
        end
      end
    end

    def matches(pattern)
      @document_content.scan(pattern).map{|match| match.first}
    end

    def unique_matches(pattern)
      matches(pattern)
    end

    alias_method :uniq_matches, :unique_matches


    def commit(new_path=nil)
      write_back_to_file(new_path)
    end

    private
    #DOCUMENT_FILE_PATH = 'word/document.xml'    
    @document_content = {}
    def read_docx_file(file)
      @document_content[file] = @zip_file.read(file)
    end

    def write_back_to_file(new_path=nil)
      if @temp_dir.nil?
        temp_file = Tempfile.new('docxedit-')
      else
        temp_file = Tempfile.new('docxedit-', @temp_dir)
      end
      Zip::OutputStream.open(temp_file.path) do |zos|
        @zip_file.entries.each do |e|
          unless @document_content.key?(e.name)
            zos.put_next_entry(e.name)
            zos.print e.get_input_stream.read
          ends
        end
        
        @document_content.each do |key, doc|
          zos.put_next_entry(key)
          zos.print val
        end
        
      end

      if new_path.nil?
        path = @zip_file.name
        FileUtils.rm(path)
      else
        path = new_path
      end
      FileUtils.mv(temp_file.path, path)
      @zip_file = Zip::File.new(path)
    end
  end
end
