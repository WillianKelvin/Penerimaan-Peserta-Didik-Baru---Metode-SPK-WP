function varargout = projek(varargin)
% PROJEK MATLAB code for projek.fig
%      PROJEK, by itself, creates a new PROJEK or raises the existing
%      singleton*.
%
%      H = PROJEK returns the handle to a new PROJEK or the handle to
%      the existing singleton*.
%
%      PROJEK('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PROJEK.M with the given input arguments.
%
%      PROJEK('Property','Value',...) creates a new PROJEK or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before projek_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to projek_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help projek

% Last Modified by GUIDE v2.5 17-May-2020 07:44:38

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @projek_OpeningFcn, ...
                   'gui_OutputFcn',  @projek_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before projek is made visible.
function projek_OpeningFcn(hObject, ~, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to projek (see VARARGIN)

% Choose default command line output for projek
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);
movegui(hObject,'center'); %biar di tengah mantap gak
% UIWAIT makes projek wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = projek_OutputFcn(~, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function nama_Callback(hObject, eventdata, handles)
% hObject    handle to nama (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of nama as text
%        str2double(get(hObject,'String')) returns contents of nama as a double


% --- Executes during object creation, after setting all properties.
function nama_CreateFcn(hObject, eventdata, handles)
% hObject    handle to nama (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function ttl_Callback(hObject, eventdata, handles)
% hObject    handle to ttl (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of ttl as text
%        str2double(get(hObject,'String')) returns contents of ttl as a double


% --- Executes during object creation, after setting all properties.
function ttl_CreateFcn(hObject, ~, ~)
% hObject    handle to ttl (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in lk.
function lk_Callback(hObject, eventdata, handles)
% hObject    handle to lk (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of lk
set(hObject,'Value',1)
set(handles.pr,'Value',0)


% --- Executes on button press in pr.
function pr_Callback(hObject, eventdata, handles)
% hObject    handle to pr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of pr
set(hObject,'Value',1)
set(handles.lk,'Value',0)


function sm1_Callback(hObject, eventdata, handles)
% hObject    handle to sm1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sm1 as text
%        str2double(get(hObject,'String')) returns contents of sm1 as a double


% --- Executes during object creation, after setting all properties.
function sm1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sm1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sm3_Callback(hObject, eventdata, handles)
% hObject    handle to sm3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sm3 as text
%        str2double(get(hObject,'String')) returns contents of sm3 as a double


% --- Executes during object creation, after setting all properties.
function sm3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sm3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sm5_Callback(hObject, eventdata, handles)
% hObject    handle to sm5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sm5 as text
%        str2double(get(hObject,'String')) returns contents of sm5 as a double


% --- Executes during object creation, after setting all properties.
function sm5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sm5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sm4_Callback(hObject, eventdata, handles)
% hObject    handle to sm4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sm4 as text
%        str2double(get(hObject,'String')) returns contents of sm4 as a double


% --- Executes during object creation, after setting all properties.
function sm4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sm4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function sm2_Callback(hObject, eventdata, handles)
% hObject    handle to sm2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of sm2 as text
%        str2double(get(hObject,'String')) returns contents of sm2 as a double


% --- Executes during object creation, after setting all properties.
function sm2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to sm2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in simpan.
function simpan_Callback(hObject, eventdata, handles)
% hObject    handle to simpan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
no = xlsread('DataSiswa.xlsx','B:B'); 

nama = get(handles.nama,'string');
ttl = get(handles.ttl,'string');
laki = get(handles.lk, 'Value');
perempuan = get(handles.pr, 'Value');
if laki == 1
    dtjk = 'Laki-laki';
elseif perempuan == 1
        dtjk = 'Perempuan';
end

 
C1 = get(handles.sm1,'String'); %memangil nilai data pada edit1
C2 = get(handles.sm2,'String'); %memangil nilai data pada edit2
C3 = get(handles.sm3,'String'); %memangil nilai data pada edit3
C4 = get(handles.sm4,'String'); %memangil nilai data pada edit4
C5 = get(handles.sm5,'String'); %memangil nilai data pada edit5
C6 = get(handles.uan,'String'); %memangil nilai data pada edit5


if numel(no) == 0 %dari no enoll
    no = 1; %nomor + 1
    range = 'A1';  %agar memulai dari rang a1
else
    no = numel(no)+1; %menambah data sesuai no + 1 intinya gitu mas susah juga jelasin
    range = strcat(['A',num2str(no)]); 
end

if isempty(C1)
    warndlg('C1 tidak boleh kosong','Warning');  %agar muncul pemberitahuan apabila tidak input data atau data kosong
elseif isempty(C2)
    warndlg('C2 tidak boleh kosong','Warning');  % sama juga
else
    data = table(cell({C1, C2, C3, C4, C5, C6, nama, ttl, dtjk}));  %memanggil data nilai pada edit-edit tadi
     writetable(data,'DataSiswa.xlsx','Sheet',1,'Range',range,'WriteVariableNames',false)  %membaca fle ditabel ,lalu menambahkan sesuai range
    msgbox(strcat(['Data No. ',num2str(no),' berhasil disimpan']),'Message') %pemberitahuan apabila sudah mengisi data
end

dataku2 = xlsread('DataSiswa.xlsx','Sheet1','A1:F10');

btC1 = get(handles.bobotC1, 'value');
if (btC1 == 1)
    bC1 = 1;
elseif (btC1 == 2)
    bC1 = 2;
elseif (btC1 == 3)
    bC1 = 3;
elseif (btC1 == 4)
    bC1 = 4;
elseif (btC1 == 5)
    bC1 = 5;
end;

btC2 = get(handles.bobotC2, 'value');
if (btC2 == 1)
    bC2 = 1;
elseif (btC2 == 2)
    bC2 = 2;
elseif (btC2 == 3)
    bC2 = 3;
elseif (btC2 == 4)
    bC2 = 4;
elseif (btC2 == 5)
    bC2 = 5;
end;

btC3 = get(handles.bobotC3, 'value');
if (btC3 == 1)
    bC3 = 1;
elseif (btC3 == 2)
    bC3 = 2;
elseif (btC3 == 3)
    bC3 = 3;
elseif (btC3 == 4)
    bC3 = 4;
elseif (btC3 == 5)
    bC3 = 5;
end;

btC4 = get(handles.bobotC4, 'value');
if (btC4 == 1)
    bC4 = 1;
elseif (btC4 == 2)
    bC4 = 2;
elseif (btC4 == 3)
    bC4 = 3;
elseif (btC4 == 4)
    bC4 = 4;
elseif (btC4 == 5)
    bC4 = 5;
end;

btC5 = get(handles.bobotC5, 'value');
if (btC5 == 1)
    bC5 = 1;
elseif (btC5 == 2)
    bC5 = 2;
elseif (btC5 == 3)
    bC5 = 3;
elseif (btC5 == 4)
    bC5 = 4;
elseif (btC5 == 5)
    bC5 = 5;
end;

btC6 = get(handles.bobotC6, 'value');
if (btC6 == 1)
    bC6 = 1;
elseif (btC6 == 2)
    bC6 = 2;
elseif (btC6 == 3)
    bC6 = 3;
elseif (btC6 == 4)
    bC6 = 4;
elseif (btC6 == 5)
    bC6 = 5;
end;


k = [0,0,0,1,0,1];
dataBobot = [bC1 bC2 bC3 bC4 bC5 bC6];
%normalisasi bobot,tahap pertama perbaikan bobot
%[m, n] = size(dataku); %inisialisasi ukuran dataku
[m, n] = size(dataku2);
dataBobot = dataBobot./sum(dataBobot); %bagi bobot per kriteria dgn jumlah total seluruh bobot


%tahapan kedua, melakukan perhitungan vektor(S) perbaris alternatif
for j = 1:n,
    if k(j)==0, dataBobot(j)=-1*dataBobot(j);
    end;
end;
for i = 1:m,
    S(i) = prod(dataku2(i,:).^dataBobot);
end;

%Menentukan nilai vektor
V = S/sum(S);
prev = transpose(V);

xlswrite('hasilpreverensi.xlsx',prev); % akan menyimpan ke excel


% --- Executes on button press in tampilkan.
function tampilkan_Callback(hObject, eventdata, handles)
% hObject    handle to tampilkan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%ReadData = xlsread('Alternatif.xlsx');
%ReadData = xlsread('DataSiswa.xlsx','Sheet1','A1:I10');
[a,b,c] = xlsread('DataSiswa.xlsx'); %[a,b,c] = 'a'untuk numeric data; 'b'untuk text data; 'c'untuk numericdan text data
set(handles.uitable2,'Data',[c]);


% --- Executes on selection change in bobotC1.
function bobotC1_Callback(hObject, eventdata, handles)
% hObject    handle to bobotC1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns bobotC1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from bobotC1


% --- Executes during object creation, after setting all properties.
function bobotC1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bobotC1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in bobotC2.
function bobotC2_Callback(hObject, eventdata, handles)
% hObject    handle to bobotC2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns bobotC2 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from bobotC2


% --- Executes during object creation, after setting all properties.
function bobotC2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bobotC2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in bobotC3.
function bobotC3_Callback(hObject, eventdata, handles)
% hObject    handle to bobotC3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns bobotC3 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from bobotC3


% --- Executes during object creation, after setting all properties.
function bobotC3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bobotC3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in bobotC4.
function bobotC4_Callback(hObject, eventdata, handles)
% hObject    handle to bobotC4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns bobotC4 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from bobotC4


% --- Executes during object creation, after setting all properties.
function bobotC4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bobotC4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in bobotC5.
function bobotC5_Callback(hObject, eventdata, handles)
% hObject    handle to bobotC5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns bobotC5 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from bobotC5


% --- Executes during object creation, after setting all properties.
function bobotC5_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bobotC5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in analisis.
function analisis_Callback(hObject, eventdata, handles)
% hObject    handle to analisis (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
%ReadData = xlsread('preverensiku.xlsx');
ReadData = xlsread('hasilpreverensi.xlsx');
set(handles.uitable1,'Data',ReadData);


% --------------------------------------------------------------------
function Untitled_1_Callback(hObject, eventdata, handles)
% hObject    handle to Untitled_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on selection change in bobotC6.
function bobotC6_Callback(hObject, eventdata, handles)
% hObject    handle to bobotC6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns bobotC6 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from bobotC6


% --- Executes during object creation, after setting all properties.
function bobotC6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to bobotC6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function uan_Callback(hObject, eventdata, handles)
% hObject    handle to uan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of uan as text
%        str2double(get(hObject,'String')) returns contents of uan as a double


% --- Executes during object creation, after setting all properties.
function uan_CreateFcn(hObject, eventdata, handles)
% hObject    handle to uan (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in hasil.
function hasil_Callback(hObject, eventdata, handles)
% hObject    handle to hasil (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

ReadData = xlsread('hasilpreverensi.xlsx');
set(handles.uitable1,'Data',ReadData);
[nilai,urutan]=max(ReadData);
set(handles.text47,'string',(nilai)); %mengeluarkan data tertinggi
set(handles.text48,'string',(urutan)); %mengeluarkan urutan keberapa



% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
set(handles.nama,'string',''); %menghapus nilai pada text4
set(handles.ttl,'string','');
set(handles.sm1,'string','');
set(handles.sm2,'string','');
set(handles.sm3,'string','');
set(handles.sm4,'string','');
set(handles.sm5,'string','');
set(handles.text47,'string','');
set(handles.text48,'string','');
set(handles.uan,'string','');
set(handles.uitable2, 'Data', {})
set(handles.uitable1, 'Data', {})
clc;
