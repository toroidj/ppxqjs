#=======================================
targetname = ppxqjs
targetext = dll
QJINCLUDE = quickjs-2025-04-26
# QuickJS のライブラリ libquickjs(.lto).a は QJLIB を参照する

ifndef X64
X64 = 1
endif
ifndef ARM
ARM = 0
endif
ifndef RELEASE
RELEASE = 0
endif

ifeq ($(strip $(ARM)),1)
  TAIL	= 64A
  QJLIB	= lib64A
else
  ifeq ($(strip $(X64)),1)
    TAIL	= 64
    QJLIB	= lib64
  else
    TAIL	= 32
    QJLIB	= lib32
  endif
endif
objdir = .obj

all:	$(objdir) code$(TAIL)$(RELEASE).o $(targetname)$(TAIL).$(targetext)
objs = $(objdir)/$(targetname).o $(objdir)/qjs_PPxObjects.o $(objdir)/qjs_CreateObject.o $(objdir)/qjs_OLEobject.o

gccname	= gcc -finput-charset=CP932 -fexec-charset=CP932
releaseopt	= -O2 -Os -flto -s
devopt	= -g
cc	= @$(gccname) $(WarnOpt) -I$(QJINCLUDE) -DQUICKJSVERSION=\"$(shell cat $(QJINCLUDE)/VERSION)\"
lddll	= @$(gccname) -shared -static $(WarnOpt)
Rcn	= wrc -l0x11 --nostdinc -l0x11 -I. -I/usr/include/wine/windows -I/usr/include/wine/wine/windows -I/usr/local/include/wine/windows

WarnOpt = -Wall -Wno-unknown-pragmas -Wl,--kill-at

link	= @$(gccname) -s
libs	= -lm -lpthread -lstdc++ -lole32 -loleaut32 -luuid


.SUFFIXES: .coff .mc .rc .mc.rc .res .res.o .idl .tlb .h  .ico .RC .c .C .cpp .CPP

$(objdir)/%.o: %.C | $(objdir)
	$(cc) -x c -c $< -o $@
$(objdir)/%.o: %.CPP | $(objdir)
	$(cc) -x c++ -c $< -o $@

# 大文字小文字を同一視する make の場合は以下２つの定義で警告が出る
$(objdir)/%.o: %.c | $(objdir)
	$(cc) -x c -c $< -o $@
$(objdir)/%.o: %.cpp | $(objdir)
	$(cc) -x c++ -c $< -o $@


.RC.coff:
	@echo $<
        ifeq ($(strip $(OldMinGW)),1)
	  $(Rcn) -fo$(basename $<).res $<
	  @windres -i $(basename $<).res -o $@
        else
	  @windres -DWINDRES -i $(basename $<).RC -o $@
        endif

.RC.res:
	@echo $<
	$(Rcn) $< -o $(basename $<).res

#------------------------------------------------------ code体系切換用
code$(TAIL)$(RELEASE).o:
	-@rm './$(targetname)$(TAIL).$(targetext)'
	-@rm '*.o'
	-@rm '$(objdir)/*.o'
	-@rm '$(objdir)/*.res'
	-@touch 'code$(TAIL)$(RELEASE).o'

$(objdir):
	-@mkdir '$(objdir)'

#------------------------------------------------------ 本体
$(targetname)$(TAIL).$(targetext): $(objs)
  ifeq ($(strip $(RELEASE)),1)
	$(lddll) $(releaseopt) $^ $(QJLIB)/libquickjs.lto.a $(libs) -o $(targetname)$(TAIL).$(targetext)
  else
	$(lddll) $(devopt) $^ $(QJLIB)/libquickjs.a $(libs) -o $(targetname)$(TAIL).$(targetext)
  endif

(objdir)/$(targetname).o: $(targetname).c ppxqjs.h
(objdir)/qjs_PPxObjects.o: qjs_PPxObjects.c ppxqjs.h
(objdir)/qjs_CreateObject.o: qjs_CreateObject.c ppxqjs.h
(objdir)/qjs_OLEobject.o: qjs_OLEobject.c ppxqjs.h
