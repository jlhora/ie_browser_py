from win32com.client import Dispatch
import time


class ie_browser:
    """Class to control an Internet Explorer Window"""

    def __init__(self, window_num=0, window_url=''):
        if window_num <= 0 and window_url == '':
            self.ie = self.__create_object()
        else:
            self.ie = self.__find_session(window_num, window_url)

    def __create_object(self):
        ie = Dispatch("InternetExplorer.Application")
        ie.Visible = 1
        return ie

    def __find_session(self, window_num=0, window_url=''):
        # CLSID for ShellWindows
        clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
        ShellWindows = Dispatch(clsid)
        if (window_url != ''):
            window_url = window_url.lower()
            for i in range(ShellWindows.Count):
                if ShellWindows[i].LocationURL.lower().find(window_url) > -1:
                    return ShellWindows[i]
            return self.__create_object()

        if (ShellWindows.Count < 1 or window_num > ShellWindows.Count or window_num <= 0):
            return self.__create_object()
        else:
            return ShellWindows[window_num - 1]

    def get_location(self):
        return self.ie.LocationURL

    def send_command(self, command):
        return eval('self.ie.Document.documentElement.' + command)

    def quit(self):
        self.ie.Quit()
        self.ie = None

    def wait_page(self):
        while self.ie.Busy:
            time.sleep(0.1)

    def busy(self):
        return self.ie.Busy

    def navigate(self, url):
        self.ie.navigate(url)
        self.wait_page()

    def get_inner_html(self):
        doc = self.ie.Document
        return doc.body.innerHTML

    def get_document(self):
        doc = self.ie.Document
        return doc

    def click_hiperlink(self, linktext):
        linktext = linktext.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'A'):
                hreftext = elem.outerText
                hreftext = hreftext.lower().strip()
                if linktext == hreftext:
                    elem.click()
                    self.wait_page()
                    return 1

        raise 'Link ' + linktext + ' was not found.'

    def wait_element(self, elem, search='id', limit=10):

        count = 0
        while count < limit:
            try:

                if search == 'id':
                    element = self.ie.get_element_by_id(elem)
                    element.focus()
                    break
                else:
                    raise "Search mode {} not available".format(search)

            except:
                pieces.time.sleep(1)
                count += 1

        if count >= limit:
            return False
        return True

    def get_element_by_id(self, elem_id='', tag='input'):
        #         print(elem_id)
        if (elem_id == ''):
            raise 'click_button(): Please specify element ID'

        doc = self.ie.Document.documentElement
        #         print(elem_id)
        elem = ''
        elem_list = doc.getElementsByTagName(tag)
        for field in elem_list:
            #             print(field.id)
            if field.id == elem_id:
                elem = field
                break

        return elem

    def click_button(self, name='', caption=''):
        if (caption == '' and name == ''):
            raise 'click_button(): Please specify either a button name or a button caption.'

        if (name != ''):
            itemtocheck = name.lower().strip()
        else:
            itemtocheck = caption.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'INPUT'):
                type = elem.getAttribute('type')
                type = type.upper()
                if (type == "IMAGE" or type == "SUBMIT" or type == "BUTTON"):
                    if name != '':
                        itemattrib = elem.getAttribute('name')
                    else:
                        itemattrib = elem.getAttribute('value')
                    itemattrib = itemattrib.lower().strip()
                    if itemattrib == itemtocheck:
                        elem.click()
                        self.wait_page()
                        return 1

        raise 'Button ' + itemtocheck + ' was not found.'

    def get_input_box(self, name):
        name = name.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'INPUT'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == name:
                    value = elem.getAttribute('value')
                    return value

        raise 'Input Element ' + name + ' was not found.'

    def set_input_box(self, name, value):
        name = name.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'INPUT'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == name:
                    elem.setAttribute('value', value)
                    return 1

        raise 'Input Element ' + name + ' was not found.'

    def get_text_area(self, name):
        name = name.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'TEXTAREA'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == name:
                    value = elem.outerText
                    return value

        raise 'Text Area ' + name + ' was not found.'

    def set_text_area(self, name, value):
        name = name.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'TEXTAREA'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == name:
                    elem.outerText = value
                    return 1

        raise 'Text Area ' + name + ' was not found.'

    def get_value_selected(self, name):
        name = name.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'SELECT'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == name:
                    index = elem.selectedIndex
                    option = elem.options(index)
                    optionvalue = option.getAttribute('value')
                    optiontext = option.innerHTML
                    return optionvalue, optiontext

        raise 'Select ' + name + ' was not found.'

    def set_value_selected(self, selname, optionvalue='', optioncaption=''):
        if (optioncaption == '' and optionvalue == ''):
            raise 'set_value_selected(): Please specify either an option value or option caption.'

        if (optionvalue != ''):
            itemtocheck = optionvalue.lower().strip()
        else:
            itemtocheck = optioncaption.lower().strip()
        selname = selname.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'SELECT'):
                elemname = elem.getAttribute('name')
                elemname = elemname.lower().strip()
                if elemname == selname:
                    for j in range(elem.length):
                        option = elem.options(j)
                        if (optionvalue != ''):
                            itemattrib = option.getAttribute('value')
                        else:
                            itemattrib = option.innerHTML
                        itemattrib = itemattrib.lower().strip()
                        if itemattrib == itemtocheck:
                            option.selected = 1
                            return 1;
                    raise 'Option ' + itemtocheck + ' was not found.'

        raise 'Select ' + selname + ' was not found.'

    def get_list_select(self, selname):
        selname = selname.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all
        listelems = []

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'SELECT'):
                elemname = elem.getAttribute('name')
                elemname = elemname.lower().strip()
                if elemname == selname:
                    for j in range(elem.length):
                        option = elem.options(j)
                        if option.selected:
                            optionvalue = option.getAttribute('value')
                            optiontext = option.innerHTML
                            listtuple = (optionvalue, optiontext)
                            listelems.append(listtuple)
                    return listelems
        raise 'List Selection ' + selname + ' was not found.'

    def get_checkbox_state(self, cbname):
        cbname = cbname.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'INPUT'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == cbname:
                    checked = elem.getAttribute('checked')
                    return checked

        raise 'Check Box ' + cbname + ' was not found.'

    def set_checkbox_state(self, cbname, checked=1):
        cbname = cbname.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'INPUT'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == cbname:
                    elem.setAttribute('checked', checked)
                    return 1

        raise 'Check Box ' + cbname + ' was not found.'

    def submit(self, formname):
        formname = formname.lower().strip()
        doc = self.ie.Document
        elemcoll = doc.all

        for i in range(elemcoll.length):
            elem = elemcoll.item(i)
            if (elem.tagName == 'FORM'):
                itemname = elem.getAttribute('name')
                itemname = itemname.lower().strip()
                if itemname == formname:
                    elem.submit()
                    self.wait_page()
                    return 1

        raise 'Form ' + formname + ' was not found.'